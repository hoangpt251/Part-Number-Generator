"""
    Chương trình tạo mã nội bộ và description
    Phát triển và kiểm thử: Phạm Trần Hoàng
    04/2026
    Tệp: display.py
"""

import os
import pandas as pd
from tkinter import messagebox

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DISPLAY_FILE = os.path.join(BASE_DIR, "core", "data", "display.xlsx")

# Đọc file nếu tồn tại, nếu không thì _sheets rỗng
_sheets = {}
if os.path.exists(DISPLAY_FILE):
    with pd.ExcelFile(DISPLAY_FILE) as xls:
        try:
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=0)
                df = df.fillna("")
                if df.shape[1] >= 2:
                    df = df.iloc[:, :2]
                    df.columns = ["display", "code"]
                    _sheets[sheet] = df
        except Exception as e:
            print(f"Lỗi đọc file display.xlsx: {e}")
            _sheets = {}
else:
    print("Không tìm thấy file display.xlsx, chỉ sử dụng entry fields")
    _sheets = {}

class DisplayLogic: 
    # Định nghĩa cấu trúc tham số chung
    # Mỗi class-code có thể có cấu trúc riêng, nhưng ở đây dùng chung
    PARAM_STRUCT = {
        "default": [
            ("Identification code (5 chars)", "entry", None),
            ("Information", "entry", None), ("Package", "entry", None)]
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
        base_fields = self.PARAM_STRUCT.get(class_code, self.PARAM_STRUCT["default"])
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
        """
            Sinh part number theo quy tắc
            P/N = 50{sub-classification}{Serial Number}
        """
        sn_no = values.get("Identification code (5 chars)", "").strip()
        
        # Đảm bảo pin number có 5 ký tự
        if len(sn_no) > 5:
            sn_no = sn_no[:5]
        
        return f"{class_code}{sn_no}R1F"

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
                messagebox.showwarning("Cảnh báo", "Kiểm tra lại ABB sub-code đã nhập, độ dài chuỗi nhỏ hơn 14 ký tự.")
                abb_processed_code = abb_code
        else: # Nếu len(ABB_code) > 15, vẫn dùng nhưng cảnh báo
            messagebox.showwarning("Cảnh báo", "Kiểm tra lại ABB sub-code đã nhập, độ dài chuỗi nhỏ hơn 14 ký tự.")
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
        """ Tạo description dạng text.
            Quy tắc: Nếu group là ABM thì sẽ yêu cầu lấy description trong bom của ABM
            Nếu group ko phải là ABM thì quy tắc đặt tên sẽ như nhau.
        """
        if group == "ABM":
            return f"Hãy copy lại description theo BOM của ABB"
        else:
            # desc. format = DSPLY {comp_type} {Device No.} 
            comp_type = self._get_component_type_name(class_code)
            info = display_values.get("Information", "")
            pkg = display_values.get("Package", "")
            
            desc_parts = ["DSPLY"]
            if comp_type:
                desc_parts.append(comp_type)
            if info:
                desc_parts.append(info)
            if pkg:
                desc_parts.append(pkg)
            desc_parts.append("ROHS")
            return " ".join(desc_parts)
        
    def _get_component_type_name(self, class_code):
        type_map = {
            "351": "LED",
            "352": "LCD",
            "353": "FLO",
            "354": "GAS",
            "355": "CRT",
            "356": "TOU",
        }
        return type_map.get(class_code, "")
    
    """Lấy giá trị hiển thị từ code"""
    def _get_display_value(self, field_name, code_value):
        sheet_name = None
        for key in self.DISPLAY_FIELDS:
            if self.DISPLAY_FIELDS.get(key) == field_name or key == field_name:
                for class_code, fields in self.PARAM_STRUCT.items():
                    for f in fields:
                        if f[0] == field_name and f[1] == "combo":
                            sheet_name = f[2]
                            break
                    if sheet_name:
                        break
                break
        
        if sheet_name and sheet_name in self.code_to_display:
            return self.code_to_display[sheet_name].get(code_value, code_value)
        return code_value

    # Hiển thi ra các parameter đã nhập
    def generate_parameter(self, class_code, values, group="Other"):
        parts = []
        # Xử lý theo group
        if group == "ABM":
            parts.append("ABM Display")
            abb = values.get("ABB Code", "")
            suffix = values.get("Suffix", "")
            suffix_display = self._get_display_value("Suffix", suffix) if suffix else suffix
            if abb:
                parts.append(f"ABB Code: {abb}")
            if suffix_display:
                parts.append(f"Suffix: {suffix_display}")        
        
        elif group == "CNN/CKR":
            parts.append("CNN/CKR Display")
            revision = values.get("Revision (1 char)", "")
            ident = values.get("CKR/CNN Code (12 chars)", "")
            if ident:
                parts.append(f"CNN/CKR Code: {ident}")
            if revision:
                parts.append(f"Revision: {revision}")
                
        else:  # group Other
            parts.append("Display")
            # Lấy giá trị hiển thị cho các combo field
            display_values = {}
            for key, val in values.items():
                if key in self.DISPLAY_FIELDS:
                    display_values[key] = self._get_display_value(key, val)
                else:
                    display_values[key] = val
            
            # Thêm các tham số còn lại
            for key, val in display_values.items():
                if val:
                    # Chuyển tên key thành dạng dễ đọc
                    display_key = key
                    if key == "Package":
                        display_key = "Package"
                    elif key == "Differ code":
                        display_key = "Differ"
                    parts.append(f"{display_key}: {val}")
            
        return ", ".join(parts)
