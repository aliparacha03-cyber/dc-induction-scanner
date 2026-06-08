import os, json, requests, openpyxl
from io import BytesIO

FILE_URL = os.environ.get('FILE_URL')
SP_TOKEN = os.environ.get('SP_TOKEN', '')

# Download the file
headers = {}
if SP_TOKEN:
    headers['Authorization'] = f'Bearer {SP_TOKEN}'

resp = requests.get(FILE_URL, headers=headers)
resp.raise_for_status()

wb = openpyxl.load_workbook(BytesIO(resp.content), data_only=True)
ws = wb.active

# Read headers from row 1
headers_row = [str(cell.value).strip() if cell.value else '' for cell in next(ws.iter_rows(min_row=1, max_row=1))]

ACTIONS = {
    "1. Top Priority - Required for Orders",
    "2. Induction from Drop Zone",
    "3. Induction from Receiving",
    "4. Induction from OS",
}

records = []

for row in ws.iter_rows(min_row=2, values_only=True):
    row_dict = {headers_row[i]: (row[i] if i < len(row) else '') for i in range(len(headers_row))}

    def get(key):
        v = row_dict.get(key, '')
        return '' if v is None else str(v).strip()

    ilpns = []
    dz = get('ASDROPZONE-iLPN ID')
    rl = get('Receiving Lane-iLPN ID')
    if dz and dz not in ('0', 'nan', 'None'):
        try:
            ilpns.append(str(int(float(dz))))
        except:
            ilpns.append(dz)
    if rl and rl not in ('0', 'nan', 'None'):
        ilpns.append(rl)

    for ilpn in ilpns:
        records.append({
            "ilpn": ilpn,
            "orgId": get('ORG ID'),
            "locationId": get('Location ID'),
            "aisle": get('Aisle'),
            "itemId": get('Item ID'),
            "locationMin": float(get('Location Min') or 0),
            "pbOnHand": float(get('PB On Hand') or 0),
            "dropzoneOnHand": float(get('ASDROPZONE-On hand') or 0),
            "osOnHand": float(get('OS On hand') or 0),
            "receivingLaneOnHand": float(get('Receiving Lane-On hand') or 0),
            "receivingLaneLocation": get('Receiving Lane-Location ID'),
            "inductionAction": get('Induction Action'),
        })

print(f"✅ Parsed {len(records)} records")

with open('data.json', 'w') as f:
    json.dump(records, f)
