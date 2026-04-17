"""
    Chương trình tạo mã nội bộ và description
    Phát triển và kiểm thử: Phạm Trần Hoàng
    04/2026
    Tệp: raw_material.py
"""

import os
import pandas as pd
from tkinter import messagebox

# Đường dẫn đến file dữ liệu của resistor
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
RAW_MATERIAL_FILE = os.path.join(BASE_DIR, "core", "data", "raw_material.xlsx")

# Đọc các sheet một lần khi import module
_sheets = {}
if os.path.exists(RAW_MATERIAL_FILE):
    with pd.ExcelFile(RAW_MATERIAL_FILE) as xls:
        try:
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=0)
                df = df.fillna("")
                if df.shape[1] >= 2:
                    df = df.iloc[:, :2]
                    df.columns = ["display", "code"]
                    _sheets[sheet] = df
        except Exception as e:
            print(f"Lỗi đọc file raw_material.xlsx: {e}")
            _sheets = {}
else:
    print("Không tìm thấy file raw_material.xlsx, chỉ sử dụng entry fields")
    _sheets = {}

class Raw_MaterialLogic:
    """
    Xử lý tham số và sinh Part Number cho các linh kiện cơ khí.
    """
    # Ánh xạ class-code -> cấu trúc tham số
    # Mỗi phần tử là tuple: (tên trường, loại, nguồn dữ liệu)
    # Loại: 'entry' (textbox), 'combo' (combobox kèm sheet name)
    PARAM_STRUCT = {
        "98": [("Sub-Classification Number (2 chars)","entry", None),
               ("Identification Code (4 chars)", "entry", None)
        ],
        "9800": [("AWG Number (2 digits)","entry", None),
                 ("Color Code", "combo", "Color")
        ],
        "99": [("Model Number (5 chars)","entry", None),
               ("Differ code (1 chars)","entry", None),
               ("Serial Number (2 chars)", "entry", None)
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
        #base_fields = self.PARAM_STRUCT.get(class_code, self.PARAM_STRUCT["default"])
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
        if class_code == "98":
            # 20{voltage-code}{capacitance value (2digits)}{tolerance-code}{package-code}{temperature-code}
            sclass_code = values.get("Sub-Classification Number (2 chars)", "")
            id_code = values.get("Identification Code (4 chars)", "").strip()
            if len(id_code) > 4:
                id_code = id_code[:4]
            return f"98{sclass_code}{id_code}R1F"
        
        elif class_code == "9800":
            # {class-code}{capacitance value (2digits)}{multiplier-code}{differ-code}
            awg_no = values.get("AWG Number (2 digits)", "")
            color_code = values.get("Color Code", "")
            return f"9800{awg_no}{color_code}R1F"
        
        elif class_code == "99":
            # {class-code}{capacitance value (2digits)}{multiplier-code}{differ-code}
            model_no = values.get("Model Number (5 chars)", "")
            difer_code = values.get("Differ code (1 chars)", "")
            sn_number = values.get("Serial Number (2 chars)", "")
            if len(model_no) > 5:
                model_no = model_no[:5]
            return f"99{model_no}{difer_code}-{sn_number}R1F"
        
        else:
            # Mặc định: ghép tất cả giá trị theo thứ tự fields
            return class_code + "".join(values.values())

    def generate_description(self, class_code, display_values, code_values, group="Other"):
        if group == "ABM":
            return f"Hãy copy lại description theo BOM của ABB"
        else:
            return f"Chờ mình ở phiên bản sau bạn nhé, hihi"
        
    def generate_parameter(self, class_code, values, group="Other"):
        return f"Đã chờ ở trên rồi thì cố chờ thêm chút bạn nhé, hihi"