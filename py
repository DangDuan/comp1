import re
import os
import datetime
from tkinter import Tk, filedialog
from openpyxl import load_workbook, Workbook

# === CLEAN FUNCTIONS ===
def clean_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = s.replace("\u00A0", "")  # non-breaking space
    s = s.replace(" ", "")
    s = s.replace("\t", "")
    s = s.replace("\r", "")
    s = s.replace("\n", "")
    return s.lower()

def similarity_ratio(str1: str, str2: str) -> float:
    str1, str2 = clean_key(str1), clean_key(str2)
    min_len = min(len(str1), len(str2))
    matches = sum(1 for i in range(min_len) if str1[i] == str2[i])
    return matches / max(len(str1), len(str2), 1)

# === MAIN ===
def update_infor(user_name="User"):
    # === chọn file bằng dialog ===
    root = Tk()
    root.withdraw()  # ẩn cửa sổ chính
    file1_path = filedialog.askopenfilename(
        title="SELECT QryTemp",
        filetypes=[("Excel Files", "*.xls *.xlsx *.xlsm")]
    )
    if not file1_path:
        print("Bạn chưa chọn file QryTemp")
        return
    
    file2_path = filedialog.askopenfilename(
        title="SELECT COMP1",
        filetypes=[("Excel Files", "*.xls *.xlsx *.xlsm")]
    )
    if not file2_path:
        print("Bạn chưa chọn file COMP1")
        return
    
    wb1 = load_workbook(file1_path, data_only=True)
    ws1 = wb1.active
    wb2 = load_workbook(file2_path)
    
    # dict key: Q(col17), value: S(col19)
    comp_dict = {}
    for row in ws1.iter_rows(min_row=2, values_only=True):
        q, s = row[16], row[18]  # Q=17, S=19
        if s:
            key = clean_key(q)
            value = str(s).strip()
            if key in comp_dict:
                if value == "Information received":
                    comp_dict[key] = value
                elif value == "Closed" and comp_dict[key] != "Information received":
                    comp_dict[key] = value
                elif value == "Waiting for Information" and comp_dict[key] not in ["Information received","Closed"]:
                    comp_dict[key] = value
            else:
                comp_dict[key] = value

    # create RESULT sheet
    if "RESULT" in wb2.sheetnames:
        wsLog = wb2["RESULT"]
    else:
        wsLog = wb2.create_sheet("RESULT")
        wsLog.append(["Sheet Name","Row","Device category","Ordering Part Number",
                      "Manufacturer name","Comment","ID","Ordering Part Number",
                      "Manufacturer name","Status","Add Comment"])

    # loop sheets in wb2
    for ws2 in wb2.worksheets:
        if ws2["B1"].value == "Component BOM Review Sheet":
            last_row = ws2.max_row
            start_row = 41
            while start_row <= last_row and ws2[f"E{start_row}"].value in (None,""):
                start_row += 1

            for i in range(start_row, last_row+1):
                key = clean_key(ws2[f"E{i}"].value)
                matchKey = None
                for dictKey in comp_dict.keys():
                    if dictKey.startswith(key):
                        matchKey = dictKey
                        break

                if matchKey:
                    ws2[f"AD{i}"] = comp_dict[matchKey]
                    ws2[f"AE{i}"] = f"{datetime.date.today()}_{user_name}\nThis components is not new type. Low risk."
                    wsLog.append([ws2.title, i, ws2[f"D{i}"].value, ws2[f"E{i}"].value,
                                  ws2[f"F{i}"].value, "MATCHED",
                                  None, matchKey, None, comp_dict[matchKey], "-"])
                else:
                    if not ws2[f"AD{i}"].value:
                        ws2[f"AD{i}"] = "Waiting for Information"
                    if not ws2[f"AE{i}"].value:
                        ws2[f"AE{i}"] = f"{datetime.date.today()}_{user_name}\nThis components is new series as Nissan.\nWe need datasheet & AEC-Q result."
                    wsLog.append([ws2.title, i, ws2[f"D{i}"].value, ws2[f"E{i}"].value,
                                  ws2[f"F{i}"].value, "Unmatched. Review again"])

    # compare second time
    for i in range(3, wsLog.max_row+1):
        if str(wsLog[f"F{i}"].value).lower() != "matched":
            unmatchedKey = clean_key(wsLog[f"D{i}"].value)
            for dictKey in comp_dict.keys():
                if similarity_ratio(unmatchedKey, dictKey) >= 0.85:
                    wsLog[f"K{i}"] = "Can was family?"
                    wsLog[f"H{i}"] = dictKey
                    wsLog[f"J{i}"] = comp_dict[dictKey]
                    break

    # fill BOM info
    for ws2 in wb2.worksheets:
        if ws2["B1"].value == "Component BOM Review Sheet":
            targetRow = 10
            while ws2[f"C{targetRow}"].value or ws2[f"E{targetRow}"].value or ws2[f"F{targetRow}"].value or ws2[f"H{targetRow}"].value:
                targetRow += 1
            ws2[f"C{targetRow}"] = datetime.date.today().strftime("%Y/%m/%d")
            if not ws2[f"E{targetRow}"].value: ws2[f"E{targetRow}"] = "Nissan"
            if not ws2[f"F{targetRow}"].value: ws2[f"F{targetRow}"] = user_name
            if not ws2[f"H{targetRow}"].value: ws2[f"H{targetRow}"] = "Updated comment"

    # save new file
    todayStr = datetime.date.today().strftime("%Y-%m-%d")
    folder = os.path.dirname(file2_path)
    base = os.path.splitext(os.path.basename(file2_path))[0]
    base = re.sub(r"(\d{4}[-/]\d{2}[-/]\d{2}|\d{8})", "", base)
    newFile = os.path.join(folder, f"{base}_{todayStr}.xlsx")
    wb2.save(newFile)
    print("✅ Done, saved as:", newFile)

# === RUN ===
if __name__ == "__main__":
    update_infor(user_name="Tester")
