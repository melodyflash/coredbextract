
import openpyxl
from openpyxl.styles import Font, Alignment, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import quote_sheetname, get_column_letter
import os
import sys
import io
import re

class ExcelBuilder:
    def __init__(self):
        self.data = {
            "Item": [],
            "Submenu": [],
            "SubmenuItem": [],
            "ModifierGroup_Items": [], # Defines Items within Groups
            "ModifierGroup": [], # Defines Group Headers (Aloha style)
            "Category": [],
            "TaxGroup": [],
            "Menu": [],
            "MenuSubmenu": []
        }

    def _get_names(self, raw_name):
        """
        Returns (ShortName, LongName).
        ShortName limited to 15 chars.
        LongName limited to 23 chars.
        """
        if not raw_name: return "", ""
        clean = str(raw_name).strip()
        long_n = clean[:23]
        short_n = clean[:15]
        return short_n, long_n

    def clean_text(self, text):
        if not text: return None
        # Allow alphanumeric, spaces, and basic punctuation
        # Remove weird hidden chars
        cleaned = re.sub(r'[^a-zA-Z0-9\s\.,\'\-\(\)\&/<> \u00C0-\u00FF]', '', str(text))
        return cleaned.strip()

    def add_data(self, json_data: dict):
        """
        Parses the AI JSON output and populates the internal data structures.
        """
        # 1. Categories - SKIPPED/CLEARED (User Request: "stop populating... category, etc")
        
        # 2. Items
        for item in json_data.get("items", []):
            try:
                price = item.get("price", 0.0)
                if isinstance(price, str):
                    price = float(price.replace('$', '').replace(',', ''))
            except:
                price = 0.0

            raw_name = item.get("name") or ""
            short_name, long_name = self._get_names(raw_name)

            row = [
                item.get("number"),
                short_name,
                long_name, 
                "Standard",  # Type (Default)
                price,
                "Item Price", # PriceMethod (Default)
                None, # TaxGroupName 
                None, # CategoryName
            ]
            
            modifiers = item.get("modifiers", [])
            modifiers += [None] * (10 - len(modifiers))
            row.extend(modifiers[:10])
            
            self.data["Item"].append(row)
            
        # 3. Modifier Group Headers + Items (Merged Logic)
        mod_item_number_start = 20000
        mod_item_count = 0

        for mg_idx, mg in enumerate(json_data.get("modifier_groups", [])):
            raw_name = mg.get("name") or ""
            mg_short, mg_long = self._get_names(raw_name)
            
            mg_num = mg.get("number")
            if not isinstance(mg_num, int):
                try: mg_num = int(mg_num)
                except: mg_num = 10000 + mg_idx
            
            if mg_num < 10000 or mg_num > 19999:
                mg_num = 10000 + (mg_idx * 10) 
            
            items = mg.get("items", [])
            rows_needed = max(len(items), 6) # Minimum 6 rows per group
            
            # Start Processing Rows
            for i in range(rows_needed):
                # Button Position Logic (7x3)
                row_pos = i // 3
                col_pos = i % 3
                
                m_item_name = None
                m_price = None 
                m_price_method = None
                
                if i < len(items):
                    m_item = items[i]
                    m_raw = m_item.get("name") or ""
                    m_short, m_long = self._get_names(m_raw)
                    
                    try: m_p = float(m_item.get("price", 0.0))
                    except: m_p = 0.0
                        
                    # Add to Item Tab (Modifiers are also Items)
                    m_number = mod_item_number_start + mod_item_count
                    mod_item_count += 1
                    self.data["Item"].append([
                        m_number, m_short, m_long, "Standard", m_p, "Item Price",
                         None, None] + [None]*10)
                    
                    m_item_name = m_long
                    m_price = "FORMULA_PRICE"
                    m_price_method = "Item Price"

                # Construct Row for ModifierGroup_Items
                # Columns: A=Number, B=ShortName, C=LongName, D=Min, E=Max, F=Free, G=Flow,
                # H=ModifierGroupName, I=ItemName, J=Price, K=Row, L=Col, M=PriceMethod
                
                if i == 0:
                    # FIRST ROW: Group Header + First Item (Merged)
                    self.data["ModifierGroup_Items"].append([
                        mg_num,           # A: Number
                        mg_short,         # B: ShortName
                        mg_long,          # C: LongName
                        None,             # D: Min
                        None,             # E: Max
                        None,             # F: Free
                        None,             # G: Flow
                        mg_long,          # H: ModifierGroupName (Ref)
                        m_item_name,      # I: ItemName
                        m_price,          # J: Price
                        row_pos,          # K: Row
                        col_pos,          # L: Col
                        m_price_method    # M: PriceMethod
                    ])
                else:
                    # SPACER ROWS
                    # Cols A-G are Blank (Protected via Logic later)
                    # "Note" goes into Col A (Number) in first spacer row (i=1)
                    
                    col_a_value = None
                    if i == 1:
                        col_a_value = "<- Right Click Number to Insert Rows"

                    self.data["ModifierGroup_Items"].append([
                        col_a_value, None, None, None, None, None, None, # A-G Blank 
                        mg_long,          # H: ModifierGroupName (Required)
                        m_item_name,      # I: ItemName
                        m_price,          # J: Price
                        row_pos,          # K: Row
                        col_pos,          # L: Col
                        m_price_method    # M: PriceMethod
                    ])

        # 4. Submenus
        for sm in json_data.get("submenus", []):
            raw_name = sm.get("name") or ""
            sm_short, sm_long = self._get_names(raw_name)
            
            self.data["Submenu"].append([
                sm.get("number"),
                sm_short,
                sm_long
            ])
            
            for idx, item_name in enumerate(sm.get("items", [])):
                row_pos = idx // 3
                col_pos = idx % 3
                
                self.data["SubmenuItem"].append([
                    sm_long,      # SubmenuName (Ref LongName)
                    "Item Button",# Type
                    item_name,    # ItemName 
                    "Item Price", # PriceMethod
                    row_pos,      # Row
                    col_pos,      # Col
                    "FORMULA_PRICE" # Price (Placeholder)
                ])

    def update_instructions_tab(self, wb):
        if "Instructions" not in wb.sheetnames: return
        ws = wb["Instructions"]
        
        # Clear specific rows (User Requests)
        for cell in ws[15]: cell.value = None
        for cell in ws[36]: cell.value = None
        # Bold Header at 26
        for cell in ws[26]: 
            if cell.value: cell.font = Font(bold=True)
        # Delete old lines
        ws.delete_rows(7, 5)
        
        # Determine TIPS location
        target_row = None
        target_col = None
        tips_content = ""
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "TIPS" in cell.value:
                    target_row = cell.row
                    target_col = cell.column
                    tips_content = cell.value
                    break
            if target_row: break
            
        if target_row:
             # Add Sheet Protection Note
             ws.cell(row=target_row-1, column=target_col, value="NOTE: To edit structure (Headers), go to Review > Unprotect Sheet. (No Password needed for basic, or '5dcr47!9').")
             
             parts = tips_content.split("TIPS")
             pre_tips = parts[0] + "TIPS"
             
             # Logic Text
             button_expl_lines = [
                pre_tips,
                "",
                "BUTTON POSITION LOGIC:",
                "1. MenuSubmenu (ButtonPositionIndex):",
                "   - Sequential number (0, 1, 2...) determining sort order.",
                "",
                "2. SubmenuItem & ModifierGroup_Items (Row/Column):",
                "   - Layout: 3 Columns x 7 Rows per page.",
                "   - Rows: Start at 0 and ascend.",
                "   - Logic: Row = Index // 3, Column = Index % 3.",
                "",
                "NOTE: Ensure no duplicate position combinations within the same Group.",
                "",
                "MODIFIER GROUPS:",
                " - First Row: Enter Group Number, Name, AND First Item.",
                " - Rows Below: Leave Group Info (Cols A-G) BLANK. Enter Items in Cols I-M.",
                " - To Add More Items: Right Click Row Number -> Insert." 
             ]
             
             # Clear old content
             max_r = ws.max_row
             if max_r >= target_row:
                 ws.delete_rows(target_row, max_r - target_row + 1)
                 
             for i, line in enumerate(button_expl_lines):
                 cell = ws.cell(row=target_row + i, column=target_col, value=line)
                 if "BUTTON POSITION LOGIC:" in line or "MODIFIER GROUPS:" in line:
                     cell.font = Font(bold=True)

    def get_template_path(self):
        filename = "Aloha_Import_Sample_3.xlsx"
        if getattr(sys, 'frozen', False):
            if hasattr(sys, '_MEIPASS'):
                path = os.path.join(sys._MEIPASS, filename)
                if os.path.exists(path): return path
            path = os.path.join(os.path.dirname(sys.executable), filename)
            if os.path.exists(path): return path
        if os.path.exists(filename): return filename
        return filename

    def build_final(self, is_empty_template=False) -> bytes:
        """
        Consolidated Logic for both Empty and Populated templates.
        """
        template_path = self.get_template_path()
        wb = openpyxl.load_workbook(template_path)
        
        self.update_instructions_tab(wb)
        
        def get_list_formula(sheet, col_range):
            return f"{quote_sheetname(sheet)}!{col_range}"

        target_sheets = ["Item", "Submenu", "SubmenuItem", "ModifierGroup_Items", "Menu", "Category", "TaxGroup", "MenuSubmenu", "ModifierGroup"]
        
        guide_notes = {
            "Item": ["Auto-ID", "Max 15 chars", "Max 23 chars (Full Name)", "Standard", "0.00", "Item Price", "Look up Tax Group", "Look up Category"],
            "SubmenuItem": ["Lookup Submenu", "Item Button", "Lookup Item", "Item Price", "0-6 (Row)", "0-2 (Col)", "=VLOOKUP_PRICE"],
            "ModifierGroup_Items": ["Auto-ID", "Max 15", "Max 23", "Min", "Max", "Free", "Flow", "Copy Name", "Lookup Item", "=VLOOKUP_PRICE", "0-6", "0-2", "Item Price"]
        }

        for sheet_name in target_sheets:
            if sheet_name not in wb.sheetnames: continue
            ws = wb[sheet_name]
            
            # 1. Clear Existing Data (Keep Header Row 1)
            if ws.max_row > 1:
                ws.delete_rows(2, amount=ws.max_row - 1)
            
            # 2. Add Guide Notes (Only to Empty Template)
            if is_empty_template:
                 notes = guide_notes.get(sheet_name)
                 if notes:
                     ws.append(notes)
                     for cell in ws[2]: cell.font = Font(italic=True, color="00808080")
            
            # 3. Handle Special Columns
            # Category: Remove OwnerName
            if sheet_name == "Category":
                owner_col = None
                for cell in ws[1]:
                     if cell.value == "OwnerName":
                         owner_col = cell.column
                         break
                if owner_col: ws.delete_cols(owner_col)

            # SubmenuItem: Add Price Header (Col G)
            if sheet_name == "SubmenuItem":
                ws.cell(row=1, column=7, value="Price")
            
            # 4. Insert Data (if not empty)
            if not is_empty_template:
                rows = self.data.get(sheet_name, [])
                if sheet_name in ["Category", "TaxGroup", "MenuSubmenu"]:
                    rows = [] # Force Clear
                    
                for idx, row_data in enumerate(rows):
                    curr_row = idx + 2 # Row 1=Header.
                    
                    # Formula Logic
                    if sheet_name == "ModifierGroup_Items":
                        # ItemName=I(9), Price=J(10)
                        if len(row_data) > 9 and row_data[9] == "FORMULA_PRICE":
                            row_data[9] = f"=IFERROR(VLOOKUP(I{curr_row}, Item!$C:$E, 3, FALSE), 0.00)"
                            
                    elif sheet_name == "SubmenuItem":
                        # ItemName=C(3), Price=G(6)
                        if len(row_data) > 6 and row_data[6] == "FORMULA_PRICE":
                            row_data[6] = f"=IFERROR(VLOOKUP(C{curr_row}, Item!$C:$E, 3, FALSE), 0.00)"
                    
                    # Clean Text
                    clean_row = []
                    for val in row_data:
                         if isinstance(val, str) and not val.startswith("="):
                             val = self.clean_text(val)
                         clean_row.append(val)
                    
                    ws.append(clean_row)
                    
                    # Style Note in ModGroup (Red Text for "<- Right Click...")
                    if sheet_name == "ModifierGroup_Items" and clean_row[0] and "Right Click" in str(clean_row[0]):
                         ws.cell(row=curr_row, column=1).font = Font(italic=True, size=9, color="00FF0000")

            # 5. Protection Setup
            # Enable Sheet Protection but Unlock Data Range
            ws.protection.password = "5dcr47!9"
            ws.protection.sheet = True
            ws.protection.insertRows = True
            ws.protection.deleteRows = True
            ws.protection.formatCells = False 
            ws.protection.formatColumns = False
            
            # Unlock large range (A2:Z1000)
            for row in ws.iter_rows(min_row=2, max_row=1500, max_col=20):
                for cell in row:
                    cell.protection = Protection(locked=False)
            
            # 6. Validations
            max_r = 1500

            if sheet_name == "Item":
                # Lookup LongName (Col C for Item, assuming Short, Long ...)
                # Actually Category usually: Num(A), Short(B), Long(C).
                # User constraint: "Dropdowns... must ALWAYS look up the longname"
                
                dv_cat = DataValidation(type="list", formula1=get_list_formula('Category', '$C$2:$C$500'), allow_blank=True) 
                ws.add_data_validation(dv_cat)
                dv_cat.add(f"H2:H{max_r}")
                
                dv_tax = DataValidation(type="list", formula1=get_list_formula('TaxGroup', '$C$2:$C$500'), allow_blank=True)
                ws.add_data_validation(dv_tax)
                dv_tax.add(f"G2:G{max_r}")
                
            elif sheet_name == "ModifierGroup_Items":
                # Lookup Item LongName (Col C)
                dv_item = DataValidation(type="list", formula1=get_list_formula('Item', '$C$2:$C$2000'), allow_blank=True)
                ws.add_data_validation(dv_item)
                dv_item.add(f"I2:I{max_r}")
                
                # *** Protection Logic using Validation ***
                # Request: "when drop down selection changes to any item... cells a through g cannot be typed into"
                # "This should only apply for the row below where the number... is established".
                # i.e. For Existing Rows 2+, If I2 is NOT Blank, A2:G2 must be Invalid/Locked.
                # Formula: `=ISBLANK($I2)`
                # Applied to A2:G1500.
                
                dv_lock = DataValidation(type="custom", formula1='=ISBLANK($I2)')
                dv_lock.error = "This row is linked to an Item. Group columns (A-G) must be blank."
                dv_lock.errorTitle = "Locked"
                ws.add_data_validation(dv_lock)
                dv_lock.add(f"A2:G{max_r}")
                
                # IMPORTANT: This might lock Row 1 entry if we pre-filled Row 1 Item?
                # User's logic: "first row of any given modifiergroup... itemname should be the first item... selected"
                # "columns a through g in this tab should only be locked once an item is selected... and keep their current state"
                # If we apply this validation to Row 2+, Row 2 (Header) has Item. So A2:G2 are locked?
                # But A2:G2 HAVE data (Group Info).
                # Validation prevents *User Entry* / Change. It does not delete existing data.
                # So if we write the data first, then apply validation, the User *cannot edit* A2:G2 if I2 is set.
                # This perfectly Matches "cannot be typed into".
                # If they want to edit Group Name, they must clear Item (I2) first?
                # Correct.
                
            elif sheet_name == "SubmenuItem":
                # Lookup LongName ($C)
                dv_sub = DataValidation(type="list", formula1=get_list_formula('Submenu', '$C$2:$C$500'), allow_blank=True)
                ws.add_data_validation(dv_sub)
                dv_sub.add(f"A2:A{max_r}")
                
                dv_item = DataValidation(type="list", formula1=get_list_formula('Item', '$C$2:$C$2000'), allow_blank=True)
                ws.add_data_validation(dv_item)
                dv_item.add(f"C2:C{max_r}")
                
                # *** Price Protection Logic ***
                # Block typing in G if D="Item Price".
                dv_price = DataValidation(type="custom", formula1=f'=D2<>"Item Price"')
                dv_price.error = "Price is automatic when 'Item Price' is selected. Change PriceMethod to edit."
                dv_price.errorTitle = "Restricted"
                ws.add_data_validation(dv_price)
                dv_price.add(f"G2:G{max_r}")

            ws.protection.enable()

        import io
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    def build_excel(self) -> bytes:
        return self.build_final(is_empty_template=False)

    def build_empty_template(self) -> bytes:
        return self.build_final(is_empty_template=True)

if __name__ == "__main__":
    print("Testing Consolidated Builder...")
    builder = ExcelBuilder()
    
    dummy_data = {
        "categories": [{"number": 1, "name": "Food"}, {"number": 2, "name": "Drinks"}],
        "items": [
            {"number": 100, "name": "Burger", "price": 10.0, "category": "Food", "modifiers": ["Sides"]},
            {"number": 101, "name": "Fries", "price": 5.0, "category": "Food"}
        ],
        "modifier_groups": [
            {"number": 1000, "name": "Sides", "min": 1, "max": 1, "items": [{"name": "Fries", "price": 0, "number":101}]}
        ],
        "submenus": [{"number": 200, "name": "Entrees", "items": ["Burger"]}]
    }
    builder.add_data(dummy_data)
    
    # 1. Test Populated Build
    data = builder.build_excel()
    with open("Aloha_Import_Template_Consolidated.xlsx", "wb") as f:
        f.write(data)
    print("Success! Created 'Aloha_Import_Template_Consolidated.xlsx'.")

    # 2. Test Empty Template Build
    empty_data = builder.build_empty_template()
    with open("Aloha_Import_Template_Consolidated_Empty.xlsx", "wb") as f:
        f.write(empty_data)
    print("Success! Created 'Aloha_Import_Template_Consolidated_Empty.xlsx'.")
