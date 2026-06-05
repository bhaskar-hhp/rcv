import html
import pandas as pd
import xml.etree.ElementTree as ET
import requests
import gspread
from google.oauth2.service_account import Credentials

SHEET_ID = '1GMPmrrePWKLcO00feVOOFwdPlVlmUI_qSAWGXDLBsY8'
SRC_GID = 92813228
DST_GID = 605373015
SA_PATH = 'rcv_service_account'

SCOPE = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
]

def get_ledger_names_from_gsheet():
    creds = Credentials.from_service_account_file(SA_PATH, scopes=SCOPE)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    ws = sh.get_worksheet_by_id(SRC_GID)
    if ws is None:
        raise ValueError(f'Worksheet with gid={SRC_GID} not found')
    col_a = ws.col_values(1)
    return [name.strip() for name in col_a if name.strip()]

print('Fetching ledger names from gsheet gid=92813228 column A...')
ledger_names = get_ledger_names_from_gsheet()

# Load existing data from destination sheet to avoid duplicates
cred = Credentials.from_service_account_file(SA_PATH, scopes=SCOPE)
gc = gspread.authorize(cred)
sh = gc.open_by_key(SHEET_ID)
dst_ws = sh.get_worksheet_by_id(DST_GID)
if dst_ws is None:
    raise ValueError(f'Worksheet with gid={DST_GID} not found')

existing = dst_ws.get_all_values()
HEADERS = ['Date', 'Ledger', 'Type', 'VoucherNo', 'DrAmt', 'CrAmt', 'LedgerName']
if existing and existing[0] and existing[0] == HEADERS:
    existing_df = pd.DataFrame(existing[1:], columns=HEADERS)
elif existing and existing[0]:
    existing_df = pd.DataFrame(existing, columns=HEADERS)
else:
    existing_df = pd.DataFrame([], columns=HEADERS)

all_new_entries = []

for ledger_name in ledger_names:
    print('\U0001f6a5', end='', flush=True)

    corrected_name = html.escape(ledger_name)

    xml_request = f'''
    <ENVELOPE>
        <HEADER>
            <VERSION>1</VERSION>
            <TALLYREQUEST>EXPORT</TALLYREQUEST>
            <TYPE>DATA</TYPE>
            <ID>LedgerVouchers</ID>
        </HEADER>
        <BODY>
            <DESC>
                <STATICVARIABLES>
                    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                    <LEDGERNAME>{corrected_name}</LEDGERNAME>
                    <EXPLODEVNUM>Yes</EXPLODEVNUM>
                </STATICVARIABLES>
                <TDL>
                    <TDLMESSAGE>
                        <REPORT Name='LedgerVouchers' ISMODIFY='Yes'>
                        </REPORT>
                    </TDLMESSAGE>
                </TDL>
            </DESC>
        </BODY>
    </ENVELOPE>
    '''

    try:
        response = requests.post('http://localhost:9002', data=xml_request)
    except Exception as e:
        print(f'Failed to contact Tally: {e}')
        continue

    try:
        root = ET.fromstring(response.text)
    except ET.ParseError as e:
        print(f'XML parsing error: {e}')
        continue

    tags = ['DSPVCHDATE', 'DSPVCHLEDACCOUNT', 'DSPVCHTYPE', 'DSPEXPLVCHNUMBER', 'DSPVCHDRAMT', 'DSPVCHCRAMT']
    buffer = []
    new_entries = []

    for elem in root.iter():
        tag = elem.tag.strip()
        text = elem.text.strip() if elem.text else ''
        if tag in tags:
            buffer.append(text)
            if len(buffer) == 6:
                entry = {
                    'Date': buffer[0],
                    'Ledger': buffer[1],
                    'Type': buffer[2],
                    'VoucherNo': buffer[5],
                    'DrAmt': buffer[3],
                    'CrAmt': buffer[4],
                    'LedgerName': ledger_name,
                }

                if not ((existing_df['Date'] == entry['Date']) & (existing_df['VoucherNo'] == entry['VoucherNo']) & (existing_df['LedgerName'] == entry['LedgerName'])).any():
                    new_entries.append(entry)
                buffer = []

    if new_entries:
        all_new_entries.extend(new_entries)

if all_new_entries:
    rows = [[e[c] for c in HEADERS] for e in all_new_entries]
    if not existing or not existing[0] or existing[0] != HEADERS:
        dst_ws.append_rows([HEADERS] + rows, value_input_option='USER_ENTERED')
    else:
        dst_ws.append_rows(rows, value_input_option='USER_ENTERED')
    print(f'\nAppended {len(all_new_entries)} new entries to gsheet gid={DST_GID}')
else:
    print('\nNo new data to append.')
