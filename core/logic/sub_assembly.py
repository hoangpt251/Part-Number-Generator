"""
    Chương trình tạo mã nội bộ và description
    Phát triển và kiểm thử: Phạm Trần Hoàng
    04/2026
    Tệp: sub_assembly.py
"""

import os
import pandas as pd
from tkinter import messagebox

# Đường dẫn đến file dữ liệu của resistor
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SUB_ASSEMBLY_FILE = os.path.join(BASE_DIR, "core", "data", "sub_assembly.xlsx")

# Đọc các sheet một lần khi import module
_sheets = {}
if os.path.exists(SUB_ASSEMBLY_FILE):
    with pd.ExcelFile(SUB_ASSEMBLY_FILE) as xls:
        try:
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=0)
                df = df.fillna("")
                if df.shape[1] >= 2:
                    df = df.iloc[:, :2]
                    df.columns = ["display", "code"]
                    _sheets[sheet] = df
        except Exception as e:
            print(f"Lỗi đọc file sub_assembly.xlsx: {e}")
            _sheets = {}
else:
    print("Không tìm thấy file sub_assembly.xlsx, chỉ sử dụng entry fields")
    _sheets = {}

class Sub_AssemblyLogic:
    """
    Xử lý tham số và sinh Part Number cho các linh kiện cơ khí.
    """
    # Ánh xạ class-code -> cấu trúc tham số
    # Mỗi phần tử là tuple: (tên trường, loại, nguồn dữ liệu)
    # Loại: 'entry' (textbox), 'combo' (combobox kèm sheet name)
    PARAM_STRUCT = {
        "91": [("Identification Number (5 chars)", "entry", None),
               ("Assembly Process Code","combo","Process"),
               ("Differ Code", "entry", None)],
        "92": [("Serial Number (6 chars)", "entry", None)],
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
        #base_fields = self.PARAM_STRUCT.get(class_code, [])
        fields = self.PARAM_STRUCT.get(class_code, [])
        """
        if group == "ABM":
            extra = self.EXTRA_PARAM["ABM"]
            fields = base_fields + extra
        elif group == "CNN/CKR":
            extra = self.EXTRA_PARAM["CNN/CKR"]
            fields = base_fields + extra
        else:
            fields = base_fields
        """

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
        
    def generate_part_number(self, class_code, values, group=None):
        """ Sinh part number theo quy tắc dựa trên class-code."""
        # Xác định loại dựa trên tiền tố class-code
        if class_code == "91":           
            id_code = values.get("Identification Number (5 chars)", "").strip()
            pr_code = values.get("Assembly Process Code", "")
            diff_code = values.get("Differ Code", "")
            # Đảm bảo độ dài 6 (có thể cắt bớt nếu dài hơn)
            if len(id_code) > 5:
                id_code = id_code[:5]
            return f"91{id_code}{pr_code}-{diff_code}R1F"
        # Nhóm 41, 42x, 43x,44x,47x,48x: serial number có 5 ký tự
        elif class_code =="92":
            sn_code = values.get("Serial Number (6 chars)", "").strip()
            if len(sn_code) > 6:
                sn_code = sn_code[:6]
            return f"92{sn_code}R1F"
        else:
            # Mặc định: ghép tất cả giá trị theo thứ tự fields
            return class_code + "".join(values.values())

    def generate_description(self, class_code, display_values, code_values, group="Other"):
            return f"Chờ mình ở phiên bản sau bạn nhé, hihi"
        
    def generate_parameter(self, class_code, values, group="Other"):
        return f"Đã chờ ở trên rồi thì cố chờ thêm chút bạn nhé, hihi"