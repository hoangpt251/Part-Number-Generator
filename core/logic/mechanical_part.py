"""
    Chương trình tạo mã nội bộ và description
    Phát triển và kiểm thử: Phạm Trần Hoàng
    04/2026
    Tệp: transistor.py
"""

import os
import pandas as pd
from tkinter import messagebox

# Đường dẫn đến file dữ liệu của resistor
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
MECHANICAL_PART_FILE = os.path.join(BASE_DIR, "core", "data", "mechanical_part.xlsx")

# Đọc các sheet một lần khi import module
_sheets = {}
if os.path.exists(MECHANICAL_PART_FILE):
    with pd.ExcelFile(MECHANICAL_PART_FILE) as xls:
        try: 
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=0)
                df = df.fillna("")
                if df.shape[1] >= 2:
                    df = df.iloc[:, :2]
                    df.columns = ["display", "code"]
                    _sheets[sheet] = df
        except Exception as e:
            print(f"Lỗi đọc file mechanical_part.xlsx: {e}")
            _sheets = {}
else:
    print("Không tìm thấy file mechanical_part.xlsx, chỉ sử dụng entry fields")
    _sheets = {}

class Mechanical_PartLogic:
    """
    Xử lý tham số và sinh Part Number cho các linh kiện cơ khí.
    """
    # Ánh xạ class-code -> cấu trúc tham số
    # Mỗi phần tử là tuple: (tên trường, loại, nguồn dữ liệu)
    # Loại: 'entry' (textbox), 'combo' (combobox kèm sheet name)
    PARAM_STRUCT = {
        "40": [("Identification Number (6 chars)", "entry", None),],
        "411": [("Serial Number (5 chars)", "entry", None),],
        "412": [("Serial Number (5 chars)", "entry", None),],
        "413": [("Serial Number (5 chars)", "entry", None),],
        "414": [("Serial Number (5 chars)", "entry", None),],
        "415": [("Serial Number (5 chars)", "entry", None),],
        "416": [("Serial Number (5 chars)", "entry", None),],
        "417": [("Serial Number (5 chars)", "entry", None),],
        "418": [("Serial Number (5 chars)", "entry", None),],
        "419": [("Serial Number (5 chars)", "entry", None),],
        "421": [("Serial Number (5 chars)", "entry", None),],
        "422": [("Serial Number (5 chars)", "entry", None),],
        "423": [("Serial Number (5 chars)", "entry", None),],
        "424": [("Serial Number (5 chars)", "entry", None),],
        "425": [("Serial Number (5 chars)", "entry", None),],
        "429": [("Serial Number (5 chars)", "entry", None),],
        "431": [("Serial Number (5 chars)", "entry", None),],
        "432": [("Serial Number (5 chars)", "entry", None),],
        "433": [("Serial Number (5 chars)", "entry", None),],
        "434": [("Serial Number (5 chars)", "entry", None),],
        "435": [("Serial Number (5 chars)", "entry", None),],
        "439": [("Serial Number (5 chars)", "entry", None),],
        "441": [("Serial Number (5 chars)", "entry", None),],
        "442": [("Serial Number (5 chars)", "entry", None),],
        "443": [("Serial Number (5 chars)", "entry", None),],
        "444": [("Serial Number (5 chars)", "entry", None),],
        "445": [("Serial Number (5 chars)", "entry", None),],
        "446": [("Serial Number (5 chars)", "entry", None),],
        "447": [("Serial Number (5 chars)", "entry", None),],
        "448": [("Serial Number (5 chars)", "entry", None),],
        "449": [("Serial Number (5 chars)", "entry", None),],
        "45": [("Identification Number (6 chars)", "entry", None),],
        "471": [("Serial Number (5 chars)", "entry", None),],
        "472": [("Serial Number (5 chars)", "entry", None),],
        "473": [("Serial Number (5 chars)", "entry", None),],
        "474": [("Serial Number (5 chars)", "entry", None),],
        "481": [("Serial Number (5 chars)", "entry", None),],
        "482": [("Serial Number (5 chars)", "entry", None),],
        "483": [("Serial Number (5 chars)", "entry", None),],
        "73": [("Identification Number (6 chars)", "entry", None),],
        "740": [("Serial Number (5 chars)", "entry", None),],
        "741": [("Serial Number (5 chars)", "entry", None),],
        "742": [("Serial Number (5 chars)", "entry", None),],
        "743": [("Serial Number (5 chars)", "entry", None),],
        "744": [("Serial Number (5 chars)", "entry", None),],
        "745": [("Serial Number (5 chars)", "entry", None),],
        "746": [("Serial Number (5 chars)", "entry", None),],
        "747": [("Serial Number (5 chars)", "entry", None),],
        "748": [("Serial Number (5 chars)", "entry", None),],
        "749": [("Serial Number (5 chars)", "entry", None),],
    }
    EXTRA_PARAM = {
        "ABM": [
            ("ABB Code", "entry", None),
            ("Suffix", "combo", "Suffix")   # Suffix sẽ được tạo trong __init__
        ],
        "CNN/CKR": [
                ("CKR/CNN Code (12 chars)", "entry", None),
                ("Revision (1 char)", "entry", None),
                ("RoHS Compliance", "combo", "ROHS"),
                ("MFG Identification (1 char)", "entry",None)
        ],
    }
    
    # Ánh xạ tên trường để lấy giá trị hiển thị từ combo
    DISPLAY_FIELDS = {
        "Package": "Package",
        "Suffix": "Suffix"
    }

    def __init__(self):
        self.sheet_data = _sheets
        # Lưu mapping display -> code và code -> display cho các sheet
        self.display_to_code = {}
        self.code_to_display = {}
        for sheet_name, df in _sheets.items():
            self.display_to_code[sheet_name] = dict(zip(df["display"], df["code"]))
            self.code_to_display[sheet_name] = dict(zip(df["code"], df["display"]))
        
        # Tạo sheet Suffix
        if "Suffix" not in self.sheet_data:
            # Tạo DataFrame cho suffix
            suffix_data = {
                "display": ["RoHS", "Green"],
                "code": ["F", "G"]
            }
            self.sheet_data["Suffix"] = pd.DataFrame(suffix_data)
            self.display_to_code["Suffix"] = {"RoHS": "F", "Green": "G"}
            self.code_to_display["Suffix"] = {"F": "RoHS", "G": "Green"}
        
        # Tạo sheet ROHS
        if "ROHS" not in self.sheet_data:
            rohs_data ={
                "display": ["RoHS", "Non-RoHS"],
                "code": ["F", "Q"]
            }
            self.sheet_data["ROHS"] = pd.DataFrame(rohs_data)
            self.display_to_code["ROHS"] = {"RoHS": "F", "Non-RoHS": "Q"}
            self.code_to_display["ROHS"] = {"F": "RoHS", "Q": "Non-RoHS"}
    
    def get_parameter_fields(self, class_code, group="Other"):
        """Trả về danh sách các trường cần nhập cho class_code"""
        # Lấy base fields từ PARAM_STRUCT (nếu có)
        base_fields = self.PARAM_STRUCT.get(class_code, [])
        if group == "ABM":
            extra = self.EXTRA_PARAM["ABM"]
            fields = base_fields + extra
        elif group == "CNN/CKR":
            extra = self.EXTRA_PARAM["CNN/CKR"]
            fields = base_fields + extra
        else:
            fields = base_fields

        # Chuyển đổi fields thành dạng có display list và code map cho combo
        result = []
        for name, typ, source in fields:
            if typ == "combo":
                df = self.sheet_data.get(source, pd.DataFrame())
                if not df.empty:
                    display_list = df["display"].tolist()
                    code_map = dict(zip(df["display"], df["code"]))
                    result.append((name, "combo", display_list, code_map))
                else:
                    # Fallback: entry nếu không có sheet
                    result.append((name, "entry", None, None))
            else:
                result.append((name, "entry", None, None))
        return result
    
    def generate_part_number(self, class_code, values, group="Other"):
        """ Sinh part number dựa trên group.
            values: dict {tên trường: giá trị đã nhập/chọn}
        """
        if group == "ABM":
            return self._generate_pn_abm(class_code, values)
        elif group == "CNN/CKR":
            return self._generate_pn_ckr(class_code, values)
        else:
            return self._generate_pn_other(class_code, values)   

    def _generate_pn_other(self, class_code, values):
        """ Sinh part number theo quy tắc dựa trên class-code."""
        # Xác định loại dựa trên tiền tố class-code
        # Nhóm 40x và 45x: serial number có 6 ký tự
        if class_code.startswith(('40','45','73')):           
            id_code = values.get("Identification Number (6 chars)", "").strip()
            # Đảm bảo độ dài 6 (có thể cắt bớt nếu dài hơn)
            if len(id_code) > 6:
                id_code = id_code[:6]
            return f"{class_code}{id_code}R1F"
        # Nhóm 41, 42x, 43x,44x,47x,48x: serial number có 5 ký tự
        elif class_code.startswith(('41','42','43','44','47','48','74')):
            sn_code = values.get("Serial Number (5 chars)", "").strip()
            if len(sn_code) > 5:
                sn_code = sn_code[:5]
            return f"{class_code}{sn_code}R1F"
        else:
            # Mặc định: ghép tất cả giá trị theo thứ tự fields
            return class_code + "".join(values.values())
        
    def _generate_pn_abm(self, class_code, values):
        class_prefix = class_code[:2] if len(class_code) >= 2 else class_code # Lấy 2 ký tự đầu của class-code
        abb_code = values.get("ABB Code", "").strip() # Xử lý ABB code
        
        # Điều kiện ràng buộc
        if len(abb_code) < 14: # Sử dụng trực tiếp    
            abb_processed_code = abb_code
        elif len(abb_code) == 14 or len(abb_code) == 15:
            if abb_code.startswith('3A'):    
                abb_processed_code = abb_code[2:] # Loại bỏ 3A và lấy phần phía sau
            else:
                # Cảnh báo nhưng vẫn dùng
                messagebox.showwarning("Cảnh báo", "Kiểm tra lại ABB code đã nhập, độ dài chuỗi nhỏ hơn 14 ký tự.")
                abb_processed_code = abb_code
        else: # Nếu len(ABB_code) > 15, vẫn dùng nhưng cảnh báo
            messagebox.showwarning("Cảnh báo", "Kiểm tra lại ABB code đã nhập, độ dài chuỗi nhỏ hơn 14 ký tự.")
            abb_processed_code = abb_code
        
        suffix = values.get("Suffix", "").strip() # Lấy suffix (có thể là code từ combo) 
        pn = f"{class_prefix}{abb_processed_code}W1{suffix}" # Tạo PN
        return pn

    """ Tạo ra part no cho CKR/CNN """
    def _generate_pn_ckr(self, class_code, values):
        class_prefix = class_code[:2] if len(class_code) >= 2 else class_code
        
        revision = values.get("Revision (1 char)", "").strip() # Lấy revision
        f_part = values.get("RoHS Compliance", "").strip() # “F” for RoHS; “Q” for Non-RoHS
        id_code = values.get("MFG Identification (1 char)", "").strip() # Manufacturer Identification
        ident_code = values.get("CKR/CNN Code (12 chars)", "").strip() 
        # Đảm bảo ident_code có 12 ký tự
        if len(ident_code) > 12:
            ident_code = ident_code[:12]
        elif len(ident_code) < 12:
            ident_code = ident_code.ljust(12, '0')           
        pn = f"{class_prefix}{ident_code}R{revision}{f_part}{id_code}"
        return pn

    def generate_description(self, class_code, display_values, code_values, group="Other"):
        if group == "ABM":
            return f"Hãy copy lại description theo BOM của ABB"
        else:
            return f"Sang phiên bản sau mình cập nhật tiếp nhé, hihi"
        
    def generate_parameter(self, class_code, values, group="Other"):
        return f"Chưa có gì update cả, hihi"
