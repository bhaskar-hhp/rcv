import requests
import xml.etree.ElementTree as ET
import csv
import json

TALLY_URL = "http://localhost:9000"
APP_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyCMZOBJC80uv2PZ76xq1IY0VR_1JA4r9bYggyUpAi3-Z6pzfIsxOSZxeEtmaF04-R3/exec"

tally_xml = """
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>EXPORT</TALLYREQUEST>
    <TYPE>COLLECTION</TYPE>
    <ID>Remote Ledger Coll</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="Remote Ledger Coll" ISINITIALIZE="Yes">
            <TYPE>Ledger</TYPE>
            <FETCH>Name,OpeningBalance,ClosingBalance</FETCH>
            <BELONGSTO>Yes</BELONGSTO>
            <ISACTIVE>Yes</ISACTIVE>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>
"""

ledger_data = []
response = requests.post(TALLY_URL, data=tally_xml.encode('utf-8'), headers={'Content-Type': 'application/xml'})

if response.status_code == 200:
    root = ET.fromstring(response.text)
    print(f"{TALLY_URL} Connected")
    for ledger in root.findall(".//LEDGER"):
        name = ledger.attrib.get("NAME", "N/A")
        opening = ledger.find("OPENINGBALANCE")
        closing = ledger.find("CLOSINGBALANCE")
        ledger_data.append({
            "name": name,
            "openingBalance": opening.text.strip() if opening is not None else "0.00",
            "closingBalance": closing.text if closing is not None else "0.00",
        })
    # Save local CSV
    with open("ledger_balances.csv", mode="w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Ledger Name", "Opening Balance", "Closing Balance"])
        for l in ledger_data:
            w.writerow([l["name"], l["openingBalance"], l["closingBalance"]])
    print(f"Saved {len(ledger_data)} ledgers to ledger_balances.csv")

    # Push to Google Sheet via Apps Script
    gs_resp = requests.post(APP_SCRIPT_URL,
        data=json.dumps({"action": "saveTallyData", "ledgers": ledger_data}),
        headers={"Content-Type": "text/plain;charset=UTF-8"})
    gs_result = gs_resp.json()
    if gs_result.get("success"):
        print(f"Pushed {gs_result['count']} ledgers to Google Sheet")
    else:
        print(f"Failed to push: {gs_result.get('error', gs_result)}")
else:
    print(f"Failed to connect to Tally. Status Code: {response.status_code}")
