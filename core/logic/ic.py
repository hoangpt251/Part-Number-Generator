"""
    Chương trình tạo mã nội bộ và description
    Phát triển và kiểm thử: Phạm Trần Hoàng
    04/2026
    Tệp: ic.py
"""

import os
import pandas as pd
from tkinter import messagebox

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
IC_FILE = os.path.join(BASE_DIR, "core", "data", "ic.xlsx")

# Đọc file nếu tồn tại, nếu không thì _sheets rỗng
_sheets = {}
if os.path.exists(IC_FILE):
    with pd.ExcelFile(IC_FILE) as xls:
        try:
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=0)
                df = df.fillna("")
                if df.shape[1] >= 2:
                    df = df.iloc[:, :2]
                    df.columns = ["display", "code"]
                    _sheets[sheet] = df
        except Exception as e:
            print(f"Lỗi đọc file ic.xlsx: {e}")
            _sheets = {}
else:
    print("Không tìm thấy file ic.xlsx, chỉ sử dụng entry fields")
    _sheets = {}

class ICLogic:
    """
    Xử lý tham số và sinh Part Number cho linh kiện IC, quy tắc:
    - Class-code bắt đầu bằng 30 hoặc 37: pn = {class-code}{identification-code (5 ký tự)}-{differ-code}
    - Class-code bắt đầu bằng 31: pn = {class-code}{identification-code (4 ký tự)}{differ-code}
    - Các class-code khác có thể định nghĩa sau
    """
    
    # Định nghĩa cấu trúc tham số chung cho tất cả IC
    # Mỗi class-code có thể có cấu trúc riêng, nhưng ở đây dùng chung
    PARAM_STRUCT = {
        # Nhóm 30x, 37x dùng chung 2 trường
        "300": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "301": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "302": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "303": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "304": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "305": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "306": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "307": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "308": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "309": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "370": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "371": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "372": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "373": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "374": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "375": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "376": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "377": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "378": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "379": [("Identification code (5 chars)", "entry", None), ("Differ code (2 chars)", "entry", None),
                ("Information", "entry", None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        # Nhóm 31x dùng 2 trường nhưng identification 4 ký tự
        "311": [("Identification code (4 chars)", "entry", None), ("Differ code (1 char)", "entry", None),
                ("Information", "entry",None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "312": [("Identification code (4 chars)", "entry", None), ("Differ code (1 char)", "entry", None),
                ("Information", "entry",None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "313": [("Identification code (4 chars)", "entry", None), ("Differ code (1 char)", "entry", None),
                ("Information", "entry",None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "314": [("Identification code (4 chars)", "entry", None), ("Differ code (1 char)", "entry", None),
                ("Information", "entry",None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "315": [("Identification code (4 chars)", "entry", None), ("Differ code (1 char)", "entry", None),
                ("Information", "entry",None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "316": [("Identification code (4 chars)", "entry", None), ("Differ code (1 char)", "entry", None),
                ("Information", "entry",None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "317": [("Identification code (4 chars)", "entry", None), ("Differ code (1 char)", "entry", None),
                ("Information", "entry",None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "318": [("Identification code (4 chars)", "entry", None), ("Differ code (1 char)", "entry", None),
                ("Information", "entry",None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        "319": [("Identification code (4 chars)", "entry", None), ("Differ code (1 char)", "entry", None),
                ("Information", "entry",None), ("Package", "entry", None), ("Number of Pin","entry",None), ("Other","entry",None)],
        # Nhóm 36:
        "36": [("Classification", "combo", "Class"), ("Identification code (4 chars)", "entry", None), 
               ("Component Type", "combo", "Comp_Type"), ("Information", "entry", None), 
               ("Package", "combo", "Package"), ("Number of Pin","entry",None), ("Other","entry",None)],
    }
    
    # Template mặc định cho các class-code chưa định nghĩa
    # Định nghĩa extra fields cho ABM và CNN/CKR
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
        "Component Type": "Component Type",
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
        # Xác định format dựa trên class_code
        if class_code.startswith(('30', '37')):
            # Nhóm 30x và 37x: identification 5 ký tự, có dấu gạch ngang trước differ
            id_code = values.get("Identification code (5 chars)", "").strip()
            diff = values.get("Differ code (2 chars)", "").strip()
            # Đảm bảo độ dài 5 (có thể cắt bớt nếu dài hơn)
            if len(id_code) > 5:
                id_code = id_code[:5]
            return f"{class_code}{id_code}-{diff}R1F"
        
        elif class_code.startswith('31'):
            # Nhóm 31x: identification 4 ký tự, không dấu gạch ngang
            id_code = values.get("Identification code (4 chars)", "").strip()
            diff = values.get("Differ code (1 char)", "").strip()
            if len(id_code) > 4:
                id_code = id_code[:4]
            return f"{class_code}{id_code}{diff}R1F"
        
        elif class_code == "36":
            # Nhóm 36: PN = {36}{identification-code}{component-type}-{package}R1F
            id_code = values.get("Identification code (4 chars)", "").strip()
            cls = values.get("Classification", "")
            package = values.get("Package", "")
            return f"36{id_code}{cls}-{package}R1F"
    
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

    """Lấy giá trị hiển thị từ code (dùng cho description)"""
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

    def generate_description(self, class_code, display_values, code_values, group="Other"):
        if group == "ABM":
            return f"Hãy copy lại description theo BOM của ABB"
        else:
            # Tạo description theo từng class-code
            if class_code.startswith(('30', '37', '31')):
                # desc. format = IC {component_type} {Device No. & revision} {package} {other function} ROHS
                comp_type = self._get_component_type_name(class_code)
            elif class_code == "36":
                comp_type = code_values.get("Component Type", "")
            
            info = display_values.get("Information", "")
            package = display_values.get("Package", "")
            no_of_pins = display_values.get("Number of Pin","")
            other = display_values.get("Other", "")
            
            desc_parts = ["IC"]
            if comp_type:
                desc_parts.append(comp_type)
            if info:
                desc_parts.append(info)
            # xử lý riêng package + pin
            if package and no_of_pins:
                desc_parts.append(f"{package}{no_of_pins}")
            else:
                if package:
                    desc_parts.append(package)
                if no_of_pins:
                    desc_parts.append(no_of_pins)
            if other:
                desc_parts.append(other)
            desc_parts.append("ROHS")
            return " ".join(desc_parts)
            
    def _get_component_type_name(self, class_code):
        type_map = {
            "300": "BIPOLAR", "301": "TTL", "302": "LS",
            "303": "L", "304": "S", "305": "H", "306": "CMOS",
            "307": "NMEM", "308": "CPU", "309": "MISC",
            "370": "AS", "371": "ASL", "372": "F", "376": "HC",
            "377": "HCT", "378": "AC/ACT",
            "311": "REG", "312": "OP/AMP", "313": "AMP",
            "314": "TIM", "315": "COM SPEC", "316": "CVRT",
            "317": "OPTO", "318": "XSTR/ARY", "319": "ANA/SPEC",
        }
        return type_map.get(class_code, "")

    # Hiển thi ra các parameter đã nhập
    def generate_parameter(self, class_code, values, group="Other"):
        parts = []
        # Xử lý theo group
        if group == "ABM":
            parts.append("ABM IC")
            abb = values.get("ABB Code", "")
            suffix = values.get("Suffix", "")
            suffix_display = self._get_display_value("Suffix", suffix) if suffix else suffix
            if abb:
                parts.append(f"ABB Code: {abb}")
            if suffix_display:
                parts.append(f"Suffix: {suffix_display}")        
        
        elif group == "CNN/CKR":
            parts.append("CNN/CKR IC")
            revision = values.get("Revision (1 char)", "")
            ident = values.get("CKR/CNN Code (12 chars)", "")
            if ident:
                parts.append(f"CNN/CKR Code: {ident}")
            if revision:
                parts.append(f"Revision: {revision}")
                
        else:  # group Other
            parts.append("IC")
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
                    if key == "Component Type":
                        display_key = "Component Type"
                    elif key == "Package":
                        display_key = "Package"
                    elif key == "Differ code":
                        display_key = "Differ"
                    parts.append(f"{display_key}: {val}")
            
        return ", ".join(parts)