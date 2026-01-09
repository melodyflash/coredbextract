
import openpyxl
from openpyxl.styles import Font, Alignment, Protection, Border, Side
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.protection import SheetProtection
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
            "ModifierGroup_Items": [], 
            "ModifierGroup": [], 
            "Category": [],
            "TaxGroup": [],
            "Menu": [],
            "MenuSubmenu": []
        }

    def _get_names(self, raw_name):
        if not raw_name: return "", ""
        clean = str(raw_name).strip()
        long_n = clean[:23]
        short_n = clean[:15]
        return short_n, long_n

    def clean_text(self, text):
        if not text: return None
        cleaned = re.sub(r'[^a-zA-Z0-9\s\.,\'\-\(\)\&/<> \u00C0-\u00FF]', '', str(text))
        return cleaned.strip()

    def add_data(self, json_data: dict):
        # 1. Categories - SKIPPED
        
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
                "Standard",  # Type
                price,
                "Item Price", # PriceMethod
                None, # TaxGroupName 
                None, # CategoryName
            ]
            
            modifiers = item.get("modifiers", [])
            modifiers += [None] * (10 - len(modifiers))
            row.extend(modifiers[:10])
            
            self.data["Item"].append(row)
            
        # 3. Modifier Group Headers + Items
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
            rows_needed = max(len(items), 6)
            
            for i in range(rows_needed):
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
                        
                    m_number = mod_item_number_start + mod_item_count
                    mod_item_count += 1
                    self.data["Item"].append([
                        m_number, m_short, m_long, "Standard", m_p, "Item Price",
                         None, None] + [None]*10)
                    
                    m_item_name = m_long
                    m_price = "FORMULA_PRICE"
                    m_price_method = "Item Price"

                if i == 0:
                    # FIRST ROW
                    self.data["ModifierGroup_Items"].append([
                        mg_num, mg_short, mg_long, None, None, None, None,
                        mg_long, m_item_name, m_price, row_pos, col_pos, m_price_method
                    ])
                else:
                    # SPACER ROWS
                    self.data["ModifierGroup_Items"].append([
                        None, None, None, None, None, None, None, 
                        mg_long, m_item_name, m_price, row_pos, col_pos, m_price_method
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
                    sm_long, "Item Button", item_name, "Item Price", row_pos, col_pos, "FORMULA_PRICE"
                ])

    def update_instructions_tab(self, wb):
        if "Instructions" not in wb.sheetnames: return
        ws = wb["Instructions"]
        
        # Formatting Fixes (User Request)
        # Rows 21, 27, 31 NOT Bold
        for r in [21, 27, 31]:
             for cell in ws[r]: cell.font = Font(bold=False)
             
        # Cell 26A BOLD
        ws["A26"].font = Font(bold=True)
        
        # Cell 36A REMOVE "TIPS"
        if ws["A36"].value and "TIPS" in str(ws["A36"].value):
             ws["A36"].value = str(ws["A36"].value).replace("TIPS", "").strip()

        # Fix: Add "5. Menu" Label at 31A
        ws["A31"].value = "5. Menu"
        # User said 31 NOT bold, but 5. Menu is a header? 
        # "rows 21, 27 and 31 do not need the entries to be bold... then make cell 26a bold"
        # I will follow instruction: 31 NOT bold.
        ws["A31"].font = Font(bold=False)
            
        # Determine TIPS location...
        target_row = None
        target_col = None
        tips_content = ""
        
        # Clean up old "TIPS" text if found elsewhere
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "TIPS" in cell.value:
                     target_row = cell.row
                     target_col = cell.column
                     tips_content = cell.value
                     # Remove the "TIPS" word effectively
                     cell.value = cell.value.replace("TIPS", "")
                     break
            if target_row: break
            
        if target_row:
             # Add Sheet Protection Note
             ws.cell(row=target_row-1, column=target_col, value="NOTE: To edit structure (Headers), go to Review > Unprotect Sheet.")
             # ... Logic text follows ...
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
        for item in (json_data.get("items") or []):
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
            
            modifiers = item.get("modifiers") or []
            modifiers += [None] * (10 - len(modifiers))
            row.extend(modifiers[:10])
            
            self.data["Item"].append(row)
            
        # 3. Modifier Group Headers + Items (Merged Logic)
        mod_item_number_start = 20000
        mod_item_count = 0

        for mg_idx, mg in enumerate(json_data.get("modifier_groups") or []):
            raw_name = mg.get("name") or ""
            mg_short, mg_long = self._get_names(raw_name)
            
            mg_num = mg.get("number")
            if not isinstance(mg_num, int):
                try: mg_num = int(mg_num)
                except: mg_num = 10000 + mg_idx
            
            if mg_num < 10000 or mg_num > 19999:
                mg_num = 10000 + (mg_idx * 10) 
            
            items = mg.get("items") or []
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
                    # Note: "Right Click" instructions now go into Comment in (Row, Col A)
                    # We add a marker to self.data to apply comment later, OR check logic in build
                    # Let's keep data clean (None) and handle comment in build loop by checking row index?
                    # Better: Put a simplified marker logic in build_final.
                    
                    self.data["ModifierGroup_Items"].append([
                        None, None, None, None, None, None, None, # A-G Blank 
                        mg_long,          # H: ModifierGroupName (Required)
                        m_item_name,      # I: ItemName
                        m_price,          # J: Price
                        row_pos,          # K: Row
                        col_pos,          # L: Col
                        m_price_method    # M: PriceMethod
                    ])

        # 4. Submenus
        for sm in (json_data.get("submenus") or []):
            raw_name = sm.get("name") or ""
            sm_short, sm_long = self._get_names(raw_name)
            
            self.data["Submenu"].append([
                sm.get("number"),
                sm_short,
                sm_long
            ])
            
            for idx, item_name in enumerate(sm.get("items") or []):
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
        
        # 1. DELETE Rows FIRST to stabilize indices
        # Deleting 5 rows starting at 7
        ws.delete_rows(7, 5)
        
        # 2. Add items 3-5 to IMPORTANT INSTRUCTIONS section (after item 2 at row 6)
        # Insert 3 new rows at row 7 for our additions
        ws.insert_rows(7, 3)
        ws.cell(row=7, column=1, value="3. All dropdowns are protected - select from the list only.")
        ws.cell(row=8, column=1, value="4. Fields left blank will be auto-generated based on database defaults.")
        ws.cell(row=9, column=1, value="5. Right-click row number -> Insert to add rows in ModifierGroup_Items.")

        # 2. FORMATTING (Targeting Final Indices)
        # User: "rows 21 27 and 31 do not need the entries to be bold"
        for r in [21, 27, 31]:
             for cell in ws[r]: 
                 if cell.value: cell.font = Font(bold=False)

        # "make cell 26a bold"
        if ws["A26"].value: ws["A26"].font = Font(bold=True)

        # "Fix: Add 5. Menu Label at 31A" (User asked for this, and says 31 not bold)
        ws["A31"].value = "5. Menu"
        ws["A31"].font = Font(bold=False)

        # 3. Clean Content ("TIPS")
        target_row = None
        target_col = None
        tips_content = ""
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "TIPS" in cell.value:
                    target_row = cell.row
                    target_col = cell.column
                    tips_content = cell.value
                    # Remove "TIPS" word
                    cell.value = cell.value.replace("TIPS", "").strip()
                    break
            if target_row: break

        if target_row:
             # Add Sheet Protection Note
             ws.cell(row=target_row-1, column=target_col, value="NOTE: To edit structure (Headers), go to Review > Unprotect Sheet.")
             
             parts = tips_content.split("TIPS")
             
             # Logic Text - No more ADDITIONAL NOTES section, it's now in the IMPORTANT INSTRUCTIONS at top
             button_expl_lines = [
                parts[0].strip(), # Pre-Tips text
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
                " - First Row: Enter Group Number, ShortName, LongName, AND First Item.",
                " - D,E,F,G (Min/Max/Free/FlowRequired): Only for group header rows.",
                " - Rows Below: Leave A-C BLANK. Enter Items in Cols I-M.",
                " - To Add More Items: Right Click Row Number -> Insert." 
             ]
             
             # Clear old content below header
             if ws.max_row >= target_row:
                 pass 
                 
             for i, line in enumerate(button_expl_lines):
                 cell = ws.cell(row=target_row + i, column=target_col, value=line)
                 if "BUTTON POSITION LOGIC:" in line or "MODIFIER GROUPS:" in line:
                     cell.font = Font(bold=True)

    def get_template_path(self):
        filename = "Aloha_Import_Template_Generated.xlsx"
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

        target_sheets = ["Item", "Submenu", "SubmenuItem", "ModifierGroup_Items", "Menu", "Category", "TaxGroup", "MenuSubmenu"]
        
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
                     for cell in ws[2]: cell.font = Font(italic=True)
            
            # 3. Handle Special Columns
            # Category: Remove OwnerName
            if sheet_name == "Category":
                owner_col = None
                for cell in ws[1]:
                     if cell.value == "OwnerName":
                         owner_col = cell.column
                         break
                if owner_col: ws.delete_cols(owner_col)

            # SubmenuItem: Add Price Header (Col G) with matching style
            if sheet_name == "SubmenuItem":
                header_cell = ws.cell(row=1, column=7, value="Price")
                # Style Match - copy formatting from adjacent header (F1)
                ref_cell = ws.cell(row=1, column=6)  # F1 (ButtonPositionColumn)
                header_cell.font = Font(name='Cambria', size=11, bold=True)
                # Copy fill and alignment from reference cell if present
                if ref_cell.fill:
                    header_cell.fill = ref_cell.fill.copy()
                if ref_cell.alignment:
                    header_cell.alignment = ref_cell.alignment.copy()
                if ref_cell.border:
                    header_cell.border = ref_cell.border.copy()
            
            # 4. Insert Data (if not empty)
            if not is_empty_template:
                rows = self.data.get(sheet_name, [])
                if sheet_name in ["Category", "TaxGroup", "MenuSubmenu"]:
                    rows = [] # Force Clear
                    
                for idx, row_data in enumerate(rows):
                    curr_row = idx + 2 # Row 1=Header.
                    
                    # Formula Logic (Price Link)
                    if sheet_name == "ModifierGroup_Items":
                        # ItemName=I(9), Price=J(10), PriceMethod=M(13)
                        # If M="Item Price", Price should be Formula.
                        # Logic: If JSON said "Item Price", we put "FORMULA_PRICE".
                        if len(row_data) > 9 and row_data[9] == "FORMULA_PRICE":
                             # Unbreakable Link attempt: The cell value is the Formula.
                             row_data[9] = f"=IFERROR(VLOOKUP(I{curr_row}, Item!$B:$E, 4, FALSE), 0.00)"
                            
                    elif sheet_name == "SubmenuItem":
                        # Item=C(3), Price=G(7) (Index 6), PriceMethod=D(4) (Index 3)
                         if len(row_data) > 6 and row_data[6] == "FORMULA_PRICE":
                             row_data[6] = f"=IFERROR(VLOOKUP(C{curr_row}, Item!$B:$E, 4, FALSE), 0.00)"
                    
                    # Clean Text
                    clean_row = []
                    for val in row_data:
                         if isinstance(val, str) and not val.startswith("="):
                             val = self.clean_text(val)
                         clean_row.append(val)
                    
                    ws.append(clean_row)
                    
                    # Style Note in ModGroup (Red Text for "<- Right Click...")
                    if sheet_name == "ModifierGroup_Items" and clean_row[0] and "Right Click" in str(clean_row[0]):
                         ws.cell(row=curr_row, column=1).font = Font(italic=True, size=9)

            # 5. Protection (Robust Setup)
            # First, completely remove any existing protection from template
            ws.protection.sheet = False

            # Now create fresh protection with NO password
            ws.protection = SheetProtection(
                sheet=True,
                formatCells=False,
                formatColumns=False,
                formatRows=False,
                insertColumns=False,
                insertRows=False,
                deleteColumns=False,
                deleteRows=False,
                sort=False,
                autoFilter=False,
                pivotTables=False,
                selectLockedCells=False,
                selectUnlockedCells=False
            )
            
            # Unlock Data Range (Row 2 - 5000)
            max_r = 5000 
            for row in ws.iter_rows(min_row=2, max_row=max_r, max_col=30):
                for cell in row:
                    cell.protection = Protection(locked=False)
            
            # 6. Validations (STRICT)
            # Helper for Strict List
            def add_strict_list(ws, formula, cell_range):
                dv = DataValidation(type="list", formula1=formula, allow_blank=True)
                dv.error = "Select from dropdown."
                dv.errorTitle = "Invalid Selection"
                dv.showErrorMessage = True # STRICT
                ws.add_data_validation(dv)
                dv.add(cell_range)

            if sheet_name == "Item":
               # Dropdowns: Type(D), Tax(G), Cat(H), ModGroups(I-R)
               # Category Name (Column B) - Import looks up by Name
               add_strict_list(ws, get_list_formula('Category', '$B$2:$B$500'), f"H2:H{max_r}")
               # TaxGroup Name (Column B) - Import looks up by Name
               add_strict_list(ws, get_list_formula('TaxGroup', '$B$2:$B$500'), f"G2:G{max_r}")

               # ModGroups (I-R) - Lookup ShortName (Column B) from ModifierGroup_Items sheet
               # Import looks up by ShortName (15 chars max)
               add_strict_list(ws, get_list_formula('ModifierGroup_Items', '$B$2:$B$200'), f"I2:R{max_r}")
               
               # Type (D) - Protected dropdown (Standard, Gift Card)
               dv_type = DataValidation(type="list", formula1='"Standard,Gift Card"', allow_blank=True)
               dv_type.showErrorMessage = True
               dv_type.error = "Select from dropdown."
               dv_type.errorTitle = "Invalid Selection"
               ws.add_data_validation(dv_type)
               dv_type.add(f"D2:D{max_r}")
               
               # PriceMethod (F) - Item Price, Price Level, etc.
               dv_pm = DataValidation(type="list", formula1='"Item Price,Price Level,Quantity Price,Ask For Price"', allow_blank=True)
               dv_pm.showErrorMessage = True
               dv_pm.error = "Select from dropdown."
               dv_pm.errorTitle = "Invalid Selection"
               ws.add_data_validation(dv_pm)
               dv_pm.add(f"F2:F{max_r}")

            elif sheet_name == "Category":
                # Type column - Protected dropdown (General, Sales, Retail)
                dv_type = DataValidation(type="list", formula1='"General,Sales,Retail"', allow_blank=True)
                dv_type.showErrorMessage = True
                dv_type.error = "Select from dropdown."
                dv_type.errorTitle = "Invalid Selection"
                ws.add_data_validation(dv_type)
                dv_type.add(f"C2:C{max_r}")  # Column C is Type in Category sheet

            elif sheet_name == "ModifierGroup_Items":
                # Lookup Item (I) - Import looks up by ShortName (Column B, 15 chars max)
                add_strict_list(ws, get_list_formula('Item', '$B$2:$B$2000'), f"I2:I{max_r}")
                
                # Price Method (M)
                dv_pm = DataValidation(type="list", formula1='"Item Price,Button Price"', allow_blank=True)
                dv_pm.showErrorMessage = True
                ws.add_data_validation(dv_pm)
                dv_pm.add(f"M2:M{max_r}")
                
                # Validation: Price (J) cannot be edited if Method="Item Price"
                # Logic: =M2<>"Item Price"
                dv_price = DataValidation(type="custom", formula1='=M2<>"Item Price"')
                dv_price.error = "Price is linked to Item Default. Change Method to edit."
                dv_price.showErrorMessage = True
                ws.add_data_validation(dv_price)
                dv_price.add(f"J2:J{max_r}")
                
                # FlowRequired (G) - Yes/No dropdown, optional (defaults to No)
                dv_flow = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
                dv_flow.showErrorMessage = True
                dv_flow.error = "Select from dropdown."
                dv_flow.errorTitle = "Invalid Selection"
                ws.add_data_validation(dv_flow)
                dv_flow.add(f"G2:G{max_r}")
                
                # Min/Max/Free Selections (D,E,F) - Require whole numbers, only when A,B,C have values
                # These are only needed for group header rows (when Number, ShortName, LongName are filled)
                dv_nums = DataValidation(
                    type="custom", 
                    formula1='=OR(ISBLANK($A2), AND(NOT(ISBLANK($A2)), NOT(ISBLANK($B2)), NOT(ISBLANK($C2))))'
                )
                dv_nums.errorTitle = "Group Row Required"
                dv_nums.error = "These fields are only for rows with Number, ShortName, and LongName filled in."
                dv_nums.showErrorMessage = True
                ws.add_data_validation(dv_nums)
                dv_nums.add(f"D2:F{max_r}")
                
                # Spacer Lock - Only for A,B,C (Group identifiers must be blank for item rows)
                # When I (ItemName) has data, A,B,C must be blank
                dv_lock = DataValidation(type="custom", formula1='=ISBLANK($I2)')
                dv_lock.error = "Number, ShortName, LongName must be blank for item rows (where ItemName is filled)."
                dv_lock.showErrorMessage = True
                ws.add_data_validation(dv_lock)
                dv_lock.add(f"A2:C{max_r}")

            elif sheet_name == "SubmenuItem":
                # Submenu (A) - Import looks up by ShortName (Column B, 15 chars max)
                add_strict_list(ws, get_list_formula('Submenu', '$B$2:$B$500'), f"A2:A{max_r}")
                # Item (C) - Import looks up by ShortName (Column B, 15 chars max)
                add_strict_list(ws, get_list_formula('Item', '$B$2:$B$2000'), f"C2:C{max_r}")
                
                # Type (B) - Protected dropdown (Item Button, PLU Button)
                dv_type = DataValidation(type="list", formula1='"Item Button,PLU Button"', allow_blank=True)
                dv_type.showErrorMessage = True
                dv_type.error = "Select from dropdown."
                dv_type.errorTitle = "Invalid Selection"
                ws.add_data_validation(dv_type)
                dv_type.add(f"B2:B{max_r}")
                
                # PriceMethod (D) - 3 options: Item Price, Button Price, Price Level
                dv_pm = DataValidation(type="list", formula1='"Item Price,Button Price,Price Level"', allow_blank=True)
                dv_pm.showErrorMessage = True
                dv_pm.error = "Select from dropdown."
                dv_pm.errorTitle = "Invalid Selection"
                ws.add_data_validation(dv_pm)
                dv_pm.add(f"D2:D{max_r}")

                # Price (G) Lock - only editable if PriceMethod is "Button Price"
                dv_p = DataValidation(type="custom", formula1='=D2<>"Item Price"')
                dv_p.error = "Price is linked to Item Default."
                dv_p.showErrorMessage = True
                ws.add_data_validation(dv_p)
                dv_p.add(f"G2:G{max_r}")

            elif sheet_name == "MenuSubmenu":
                # MenuName (A) - Import looks up by ShortName (Column B, 15 chars max)
                add_strict_list(ws, get_list_formula('Menu', '$B$2:$B$100'), f"A2:A{max_r}")
                # SubmenuName (B) - Import looks up by ShortName (Column B, 15 chars max)
                add_strict_list(ws, get_list_formula('Submenu', '$B$2:$B$500'), f"B2:B{max_r}")


            # Protection is already enabled via ws.protection.sheet = True above
            # Do NOT call ws.protection.enable() as it resets all permission flags


        # Apply Logic for Comment (Post-Append)
        if sheet_name == "ModifierGroup_Items":
            # Iterate to find "Right Click" text in Col A?
            # Actually, in `add_data` let's pass it.
            # If we detect it in data, we remove from value and add comment.
            for row_idx in range(2, ws.max_row + 1): # Start from row 2
                cell = ws.cell(row=row_idx, column=1)
                # Check if it's a spacer row (Col A is blank, but Col H (ModifierGroupName) is not)
                # And if it's the second row of a group (i.e., the first spacer row)
                # This is tricky without tracking group state.
                # Simpler: Add comment to any row where A is blank and H is not, and it's not the first row of a group.
                # The original logic put "<- Right Click Number to Insert Rows" in A for i=1 (second row of group).
                # Since add_data now puts None, we need a different trigger.
                # Let's assume the comment should be on the first blank 'Number' cell (A) for each group.
                # This means if A is blank, and the previous row's A was not blank (start of a new group).
                
                # This logic is complex to implement robustly here without more context from add_data.
                # Sticking to the provided edit's comment logic:
                # "If we detect it in data, we remove from value and add comment."
                # But add_data now puts None.
                # The instruction says "Update add_data to use Comment instead of text note."
                # And the code edit for add_data removes the text note.
                # The comment in build_final says "If we detect it in data, we remove from value and add comment."
                # This is contradictory.
                # I will follow the explicit code edit for add_data (remove text) and the comment logic in build_final
                # which implies a marker.
                # Since the marker is removed from add_data, I will add a simple comment to the first blank A cell
                # after a non-blank A cell in ModifierGroup_Items.
                
                # Let's re-evaluate the original intent:
                # "Note" goes into Col A (Number) in first spacer row (i=1)
                # The `add_data` change removes `col_a_value = "<- Right Click Number to Insert Rows"`
                # So `clean_row[0]` will always be `None` for spacer rows.
                # The `build_final` comment logic `if cell.value and "Right Click" in str(cell.value):` will never trigger.
                
                # The instruction is "Update add_data to use Comment instead of text note."
                # This implies the comment should be created *when the data is appended*.
                # However, openpyxl.Comment objects cannot be directly stored in `self.data`.
                # They must be applied to a cell *after* the cell is created in the worksheet.
                
                # Given the `add_data` change to `None` for A-G, and the `build_final` comment:
                # "Let's keep data clean (None) and handle comment in build loop by checking row index?
                # Better: Put a simplified marker logic in build_final."
                # This suggests `build_final` should infer where to put the comment.
                
                # Let's implement the inference:
                # A comment "Right Click Number to Insert Rows" should be placed on cell A of the *first* spacer row
                # for each modifier group.
                # A spacer row is identified by A being blank, but H (ModifierGroupName) being present.
                # The *first* spacer row is the one immediately following a row where A (Number) was not blank.
                
                if row_idx > 2: # Cannot be the first data row (row 2)
                    current_cell_A = ws.cell(row=row_idx, column=1)
                    prev_cell_A = ws.cell(row=row_idx - 1, column=1)
                    current_cell_H = ws.cell(row=row_idx, column=8) # ModifierGroupName
                    
                    if current_cell_A.value is None and prev_cell_A.value is not None and current_cell_H.value is not None:
                        current_cell_A.comment = Comment("Right Click Number to Insert Rows", "System")

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
