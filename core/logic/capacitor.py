"""
    Chương trình tạo mã nội bộ và description
    Phát triển và kiểm thử: Phạm Trần Hoàng
    03/2026
    Tệp: capacitor.py
"""

import os
import pandas as pd
from tkinter import messagebox

# Đường dẫn đến file dữ liệu của Capacitor
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
CAPACITOR_FILE = os.path.join(BASE_DIR, "core", "data", "capacitor.xlsx")

# Đọc các sheet một lần khi import module
_sheets = {}
if os.path.exists(CAPACITOR_FILE):
    with pd.ExcelFile(CAPACITOR_FILE) as xls:
        try:
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=0)
                df = df.fillna("")
                if df.shape[1] >= 2:
                    df = df.iloc[:, :2]
                    df.columns = ["display", "code"]
                    _sheets[sheet] = df
                    #print(f"Đã tải sheet: {sheet}")
        except Exception as e:
            print(f"Lỗi đọc file capacitor.xlsx: {e}")
            _sheets = {}
else:
    print("Không tìm thấy file capacitor.xlsx, chỉ sử dụng entry fields")
    _sheets = {}

class CapacitorLogic:
    """ Xử lý tham số và sinh Part Number cho linh kiện Capacitor 
        Ánh xạ class-code -> cấu trúc tham số
        Mỗi phần tử là tuple: (tên trường, loại, nguồn dữ liệu)
        Loại: 'entry' (textbox), 'combo' (combobox kèm sheet name)
    """
    PARAM_STRUCT = {
        "20": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (3 chars)", "entry", None),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code")
        ],
        "211": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "212": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "221": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "222": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "223": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "224": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "225": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "226": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "227": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "230": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "231": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "232": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "233": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "240": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "241": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "242": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "243": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "250": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "251": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "252": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "253": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "260": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "261": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "262": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "263": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],       
        "281": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
        "282": [
            ("Voltage", "combo", "Voltage"),
            ("Capacitance (2 chars)", "entry", None),
            ("Multiplier", "combo", "Multiplier"),
            ("Tolerance", "combo", "Tolerance"),
            ("Package", "combo", "Package"),
            ("Temperature coefficient", "combo", "Temp_Code"),
            ("Differ code (2 chars)", "entry", None)
        ],
    }

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

    # Ánh xạ tên trường để lấy giá trị hiển thị từ combo => phải sửa cho phù hợp với tụ điện
    DISPLAY_FIELDS = {
        "Voltage": "Voltage",
        "Multiplier": "Multiplier",
        "Tolerance": "Tolerance",
        "Package": "Package",
        "Temperature": "Temperature",
        "Suffix": "Suffix"
    }

    MULTIPLIER_MAP = {
        "10^0": 1,
        "10^1": 10,
        "10^2": 100,
        "10^3": 1000,
        "10^4": 10000,
        "10^5": 100000,
        "10^6": 1000000,
        "10^7": 10000000,
        "10^8": 100000000,
        "10^9": 1000000000,
        "10^-1": 0.1,
        "10^-2": 0.01,
        "10^-3": 0.001,
        "10^-4": 0.0001
    }

    def __init__(self):
        self.sheet_data = _sheets # Lưu mapping giữa tên trường và sheet dữ liệu (cho combo)

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
        """ Trả về danh sách các trường cần nhập.
            - Default: lấy từ PARAM_STRUCT
            - ABM: lấy từ PARAM_STRUCT (nếu có) + thêm ABB Code và Suffix
            - CNN/CKR: lấy từ PARAM_STRUCT (nếu có) + thêm các trường đặc biệt
        """
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
        if class_code == "20":
            # 20{voltage-code}{capacitance value (2digits)}{tolerance-code}{package-code}{temperature-code}
            voltage = values.get("Voltage", "")
            cap_val = values.get("Capacitance (3 chars)", "")
            tol = values.get("Tolerance", "")
            pkg = values.get("Package", "")
            temp = values.get("Temperature coefficient", "")
            return f"20{voltage}{cap_val}{tol}{pkg}-{temp}R1F"
        elif class_code in ("211", "212", "221", "222", "223", "224", "225", "226", "227","281","282"):
            # {class-code}{capacitance value (2digits)}{multiplier-code}{differ-code}
            cap_val = values.get("Capacitance (2 chars)", "")
            mult = values.get("Multiplier", "")
            diff = values.get("Differ code (2 chars)", "")
            return f"{class_code}{cap_val}{mult}{diff}R1F"
        elif class_code in ("230", "231", "232", "233", "240", "241", "242", "243", "250", "251", "252", "253", "260", "261", "262", "263"):
            # {class-code}{capacitance value (2digits)}{multiplier-code}{tolerance-code}{voltage}{differ-code}
            cap_val = values.get("Capacitance (2 chars)", "")
            mult = values.get("Multiplier", "")
            tol = values.get("Tolerance", "")
            voltage = values.get("Voltage", "")
            diff = values.get("Differ code (2 chars)", "")
            return f"{class_code}{cap_val}{mult}{tol}{voltage}-{diff}R1F"
        else:
            # Mặc định: ghép tất cả giá trị theo thứ tự fields
            return class_code + "".join(values.values())

    """ Tạo ra part no cho ABM """
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

    def _format_capacitance_value(self, capacitance_value, multiplier_code=None):
        """
        Định dạng giá trị điện dung.
        - Nếu capacitance_value có 3 ký tự: value = 2 số đầu * 10^(số cuối)
        - Nếu có 2 ký tự (hoặc khác): value = capacitance_value * multiplier (nếu có)
        Trả về chuỗi đã định dạng với đơn vị PF, NF, UF, MF, F.
        """
        try:
            # Xác định giá trị thực
            if len(capacitance_value) == 3:
                # Dạng 3 số: 2 số đầu * 10^(số cuối)
                first_two = float(capacitance_value[:2])
                last_one = float(capacitance_value[2])
                real_value = first_two * (10 ** last_one)
            else:
                # Dạng 2 số hoặc khác: nhân với multiplier nếu có
                cap_val = float(capacitance_value) if capacitance_value else 0
                mult_value = 1
                if multiplier_code:
                    # Lấy giá trị hiển thị của multiplier từ code (nếu cần)
                    multiplier_display = self.code_to_display.get("Multiplier", {}).get(multiplier_code, multiplier_code)
                    mult_value = self.MULTIPLIER_MAP.get(multiplier_display, 1)
                real_value = cap_val * mult_value

            # Định dạng đơn vị
            if real_value >= 1_000_000_000_000:
                value = real_value / 1_000_000_000_000
                unit = "F"
            elif real_value >= 1_000_000_000:
                value = real_value / 1_000_000_000
                unit = "MF"
            elif real_value >= 1_000_000:
                value = real_value / 1_000_000
                unit = "UF"
            elif real_value >= 1_000:
                value = real_value / 1_000
                unit = "NF"
            else:
                value = real_value
                unit = "PF"

            # Làm tròn và loại bỏ .0 nếu cần
            if value.is_integer():
                return f"{int(value)}{unit}"
            else:
                # Giữ 1 chữ số thập phân, bỏ .0 nếu có
                formatted = f"{value:.1f}{unit}".rstrip('0').rstrip('.')
                return formatted if formatted.endswith(unit) else formatted + unit

        except (ValueError, TypeError):
            # Fallback: trả về giá trị gốc
            return f"{capacitance_value}{multiplier_code or ''}"

    def generate_description(self, class_code, display_values, code_values, group="Other"):
        """ Tạo description dạng text.
            Quy tắc: Nếu group là ABM thì sẽ yêu cầu lấy description trong bom của ABM
            Nếu group ko phải là ABM thì quy tắc đặt tên sẽ như nhau.
        """
        if group == "ABM":
            return f"Hãy copy lại description theo BOM của ABB"
        else:
            # Xử lý riêng cho capacitance value và multiplier để có giá trị thực
            if "Capacitance (2 chars)" in display_values and "Multiplier" in display_values:
                capacitance_raw = display_values["Capacitance (2 chars)"]
                multiplier_code = display_values.get("Multiplier", "")
                display_values["Capacitance formatted"] = self._format_capacitance_value(capacitance_raw, multiplier_code)
            elif "Capacitance (3 chars)" in display_values:
                capacitance_raw = display_values["Capacitance (3 chars)"]
                display_values["Capacitance formatted"] = self._format_capacitance_value(capacitance_raw)

            # Tạo description theo từng class-code
            # desc. format = CAP aaa bbbV ccc eee% fff ggg
            
            aaa = self._get_component_type_name(class_code) # capacitor type
            bbb = display_values.get("Voltage", "")         # voltage
            ccc = display_values.get("Capacitance formatted", "")     # Capacitance value
            eee = display_values.get("Tolerance", "")               # tolerance
            fff = display_values.get("Package", "")                 # package
            ggg = display_values.get("Temperature coefficient", "") # Temperature coefficient
            
            desc_parts = ["CAP"]
            if aaa:
                desc_parts.append(aaa)
            if bbb:
                if bbb == "Others":
                    desc_parts.append("")
                else:
                    #bbb1 = bbb[:-3]
                    bbb1 = bbb.replace("VDC", "")
                    desc_parts.append(f"{bbb1}V")
            if ccc:
                desc_parts.append(ccc)
            if eee:
                desc_parts.append(eee)
            if fff:
                desc_parts.append(fff)
            if ggg:
                desc_parts.append(ggg)
            desc_parts.append("ROHS")
            return " ".join(desc_parts)
    
    def _get_component_type_name(self, class_code):
        type_map = {
            "20": "CER",
            "211": "MCA", "212": "MCA",
            "221": "MYL",
            "222": "PC",
            "223": "MPC",
            "224": "PE",
            "225": "MPE",
            "226": "PS",
            "227": "PP",
            "230": "TAN", "231": "TAN", "232": "TAN", "233": "TAN",
            "240": "ALU", "241": "ALU", "242": "ALU", "243": "ALU",
            "260": "NET", "261": "NET", "262": "NET", "263": "NET",
            "281": "VAR", "282": "VAR",
        }         
        return type_map.get(class_code, "")  

    # Hiển thi ra các parameter đã nhập
    def generate_parameter(self, class_code, values, group="Other"):
        parts = []
        # Xử lý theo group
        if group == "ABM":
            parts.append("ABM Capacitor")
            abb = values.get("ABB Code", "")
            suffix = values.get("Suffix", "")
            suffix_display = self._get_display_value("Suffix", suffix) if suffix else suffix
            if abb:
                parts.append(f"ABB Code: {abb}")
            if suffix_display:
                parts.append(f"Suffix: {suffix_display}")        
        
        elif group == "CNN/CKR":
            parts.append("CNN/CKR Capacitor")
            revision = values.get("Revision (1 char)", "")
            ident = values.get("CKR/CNN Code (12 chars)", "")
            if ident:
                parts.append(f"CNN/CKR Code: {ident}")
            if revision:
                parts.append(f"Revision: {revision}")
                
        else:  # group Other
            parts.append("Capacitor")
            # Lấy giá trị hiển thị cho các combo field
            display_values = {}
            for key, val in values.items():
                if key in self.DISPLAY_FIELDS:
                    display_values[key] = self._get_display_value(key, val)
                else:
                    display_values[key] = val
            
            # Xử lý capacitance value và multiplier
            if "Capacitance (2 chars)" in display_values and "Multiplier" in values:
                capacitance_raw = display_values.get("Capacitance (2 chars)", "")
                multiplier_code = values.get("Multiplier", "")
                if capacitance_raw and multiplier_code:
                    capacitance_formatted = self._format_capacitance_value(capacitance_raw, multiplier_code)
                    parts.append(f"Capacitance: {capacitance_formatted}")
                    # Loại bỏ các key đã xử lý để tránh lặp
                    display_values.pop("Capacitance (2 chars)", None)
                    # Multiplier không được thêm trực tiếp
            elif "Capacitance (3 chars)" in display_values and "Multiplier" in values:
                capacitance_raw = display_values.get("Capacitance (3 chars)", "")
                multiplier_code = values.get("Multiplier", "")
                if capacitance_raw and multiplier_code:
                    capacitance_formatted = self._format_capacitance_value(capacitance_raw, multiplier_code)
                    parts.append(f"Capacitance: {capacitance_formatted}")
                    # Loại bỏ các key đã xử lý để tránh lặp
                    display_values.pop("Capacitance (3 chars)", None)
            else:
                # Nếu không có multiplier, vẫn hiển thị capacitance value
                if "Capacitance (2 chars)" in display_values and display_values["Capacitance (2 chars)"]:
                    parts.append(f"Capacitance (2 chars): {display_values['Capacitance (2 chars)']} farad")
                    display_values.pop("Capacitance (2 chars)", None)
                elif "Capacitance (3 chars)" in display_values and display_values["Capacitance (3 chars)"]:
                    parts.append(f"Capacitance (3 chars): {display_values['Capacitance (3 chars)']} farad")
                    display_values.pop("Capacitance (3 chars)", None)
            
            # Thêm các tham số còn lại
            for key, val in display_values.items():
                if val:
                    # Chuyển tên key thành dạng dễ đọc
                    display_key = key
                    if key == "Voltage":
                        display_key = "Voltage"
                    if key == "Multiplier":
                        display_key = "Multiplier"
                    elif key == "Tolerance":
                        display_key = "Tolerance"
                    elif key == "Package":
                        display_key = "Package"
                    elif key == "Temperature":
                        display_key = "Temperature"
                    elif key == "Differ code":
                        display_key = "Differ"
                    parts.append(f"{display_key}: {val}")
        
        return ", ".join(parts)