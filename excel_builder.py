import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import quote_sheetname
import io
import shutil
import os
import sys

class ExcelBuilder:
    """
    Modifies the existing Aloha_Import_Sample.xlsx template.
    Populates it with extracted data while preserving layout and validations.
    """
    
    # Map headers to list index for easy data insertion
    SCHEMA_MAP = {
        "Category": ["Number", "Name", "Type", "MaxPerCheck", "Description", "OwnerName"],
        "TaxGroup": ["Number", "Name"],
        "Item": [
            "Number", "ShortName", "LongName", "Type", "DefaultPrice", "PriceMethod",
            "TaxGroupName", "CategoryName", "ModifierGroup1", "ModifierGroup2",
            "ModifierGroup3", "ModifierGroup4", "ModifierGroup5", "ModifierGroup6",
            "ModifierGroup7", "ModifierGroup8", "ModifierGroup9", "ModifierGroup10"
        ],
        "ModifierGroup_Items": [
            "Number", "ShortName", "LongName", "MinimumSelections", "MaximumSelections",
            "FreeSelections", "FlowRequired", "ModifierGroupName", "ItemName",
            "Price", "ButtonPositionRow", "ButtonPositionColumn", "PriceMethod"
        ],
        "Menu": ["Number", "ShortName", "LongName"],
        "Submenu": ["Number", "ShortName", "LongName"],
        "MenuSubmenu": ["MenuName", "SubmenuName", "ButtonPositionIndex"],
        "SubmenuItem": [
            "SubmenuName", "Type", "ItemName", "PriceMethod",
            "ButtonPositionRow", "ButtonPositionColumn"
        ]
    }

    def __init__(self):
        # We don't load here, we load during build to ensure fresh copy
        self.data = {sheet: [] for sheet in self.SCHEMA_MAP.keys()}

    def add_data(self, json_data: dict):
        """
        Parses the AI JSON output and populates the internal data structures.
        """
        # 1. Categories - SKIPPED (User wants to manage manually / always submenus)
        # self.data["Category"] = [] 
            
        # 2. Items
        # Map item names to numbers for reference
        item_map = {}
        
        for item in json_data.get("items", []):
            try:
                # Handle price parsing safely
                price = item.get("price", 0.0)
                if isinstance(price, str):
                    price = float(price.replace('$', '').replace(',', ''))
            except:
                price = 0.0

            name = item.get("name") or ""
            short_name = name[:15]
            description = item.get("description")
            # If description is None or empty, fallback to name
            if not description:
                description = name
            long_name = description[:23]

            row = [
                item.get("number"),
                short_name,
                long_name, 
                None, # Type (Blank)
                price,
                None, # PriceMethod (Blank)
                None, # TaxGroupName (Blank)
                None, # CategoryName (Blank)
            ]
            
            modifiers = item.get("modifiers", [])
            modifiers += [None] * (10 - len(modifiers))
            row.extend(modifiers[:10])
            
            self.data["Item"].append(row)
            item_map[short_name] = item.get("number")
            
        # 3. Modifier Group Headers (Titles Only)
        for mg_idx, mg in enumerate(json_data.get("modifier_groups", [])):
            name = mg.get("name") or ""
            short_name = name[:15]
            long_name = name[:23]
            
            # Enforce constraints: Number between 10000 and 19999 for ModifierGroup
            mg_num = mg.get("number")
            if not isinstance(mg_num, int):
                try:
                    mg_num = int(mg_num)
                except:
                    mg_num = 10000 + mg_idx
            
            if mg_num < 10000 or mg_num > 19999:
                # Force into range. 
                # If it's too small, add 10000. 
                # If still out of range or collision likely, just generate strictly.
                mg_num = 10000 + (mg_idx * 10) # Safe default generation
            
            # ONLY Number, ShortName, LongName, ModifierGroupName(Same as ShortName)
            self.data["ModifierGroup_Items"].append([
                mg_num,
                short_name,
                long_name, 
                None, # Min
                None, # Max
                None, # Free
                None, # Flow
                short_name, # ModifierGroupName (Same as ShortName)
                None, # ItemName
                None, # Price
                None, # Row
                None, # Col
                None  # PriceMethod
            ])

        # 4. Submenus
        submenu_map = {}
        for sm in json_data.get("submenus", []):
            self.data["Submenu"].append([
                sm.get("number"),
                sm.get("name"),
                sm.get("name")
            ])
            
            # Add items to SubmenuItem
            # We need to look up item numbers or just use names? 
            # SubmenuItem schema: SubmenuName, Type, ItemName...
            for idx, item_name in enumerate(sm.get("items", [])):
                self.data["SubmenuItem"].append([
                    sm.get("name"),
                    None, # Type (Blank)
                    item_name,
                    None, # PriceMethod (Blank)
                    (idx // 4) + 1, # Row (1-based?)
                    (idx % 4) + 1  # Col (1-based?)
                ])
                
    def get_template_path(self):
        """Resolves the path to the template file, handling frozen state (PyInstaller)."""
        filename = "Aloha_Import_Sample.xlsx"
        if getattr(sys, 'frozen', False):
            # If frozen, look in sys._MEIPASS (onefile) or relative to executable (onedir)
            if hasattr(sys, '_MEIPASS'):
                path = os.path.join(sys._MEIPASS, filename)
                if os.path.exists(path):
                    return path
            
            # Fallback for onedir or if not in MEIPASS
            path = os.path.join(os.path.dirname(sys.executable), filename)
            if os.path.exists(path):
                return path
        
        # Default local path
        return filename

    def build_excel(self) -> bytes:
        """
        Loads the template, clears old data, inserts new data, adds validations, returns bytes.
        """
        # Load Template
        template_path = self.get_template_path()
        try:
            wb = openpyxl.load_workbook(template_path)
        except FileNotFoundError:
            raise FileNotFoundError(f"Template file '{template_path}' not found!")

        for sheet_name, rows in self.data.items():
            if sheet_name not in wb.sheetnames:
                continue
                
            ws = wb[sheet_name]
            
            # Clear existing data Logic
            # User wants to preserve static data (Categories/Tax) if we aren't extracting it.
            # We only clear "transactional" sheets: Item, Submenu, SubmenuItem, ModifierGroup_Items (if valid)
            # Actually, user wants ModGroup titles. If we clear, we lose existing. 
            # If we don't clear, we might duplicate. 
            # Let's clear sheets we are actively populating, but maybe NOT Category/TaxGroup?
            should_clear = sheet_name in ["Item", "Submenu", "SubmenuItem", "ModifierGroup_Items", "Menu"]
            
            if should_clear and ws.max_row > 1:
                ws.delete_rows(2, amount=ws.max_row-1)
                
            # Insert New Data
            # Note: For ModifierGroup_Items, if we cleared, we just add our headers.
            for row_data in rows:
                ws.append(row_data)
                
            # Apply Validations
            max_row = ws.max_row
            if max_row < 2: max_row = 100 
            
            # Helper to create safe validation formula (removes manual single quotes logic collision)
            # quote_sheetname handles necessary quoting.
            def get_list_formula(sheet, col_range):
                return f"{quote_sheetname(sheet)}!{col_range}"

            if sheet_name == "Item":
                # Category (Col H / 8) -> Category!B
                # Ensure validation points to valid range even if empty.
                dv_cat = DataValidation(type="list", formula1=get_list_formula('Category', '$B$2:$B$1000'), allow_blank=True)
                ws.add_data_validation(dv_cat)
                dv_cat.add(f"H2:H{max_row+500}")
                
                # TaxGroup (Col G / 7) -> TaxGroup!B
                dv_tax = DataValidation(type="list", formula1=get_list_formula('TaxGroup', '$B$2:$B$1000'), allow_blank=True)
                ws.add_data_validation(dv_tax)
                dv_tax.add(f"G2:G{max_row+500}")
                
                # ModifierGroups 1-10 (Col I to R / 9-18) -> ModifierGroup_Items!B
                # User said "titles of modifier groups". We assume they are in Col B of ModifierGroup_Items.
                dv_mod = DataValidation(type="list", formula1=get_list_formula('ModifierGroup_Items', '$B$2:$B$2000'), allow_blank=True)
                ws.add_data_validation(dv_mod)
                # Apply to columns I through R
                for col_char in ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']:
                     dv_mod.add(f"{col_char}2:{col_char}{max_row+500}")
            
            elif sheet_name == "ModifierGroup_Items":
                # ItemName (Col I / 9) -> Item!B
                dv_item = DataValidation(type="list", formula1=get_list_formula('Item', '$B$2:$B$2000'), allow_blank=True)
                ws.add_data_validation(dv_item)
                dv_item.add(f"I2:I{max_row+500}")
                
                # ModifierGroupName (Col H / 8) Formula
                for r in range(2, max_row + 1):
                    # Only apply if row has data (we might have appended blank rows?)
                     if ws[f"B{r}"].value:
                        ws[f"H{r}"] = f"=B{r}"
                    
            elif sheet_name == "MenuSubmenu":
                # Menu (A) -> Menu!B
                dv_menu = DataValidation(type="list", formula1=get_list_formula('Menu', '$B$2:$B$500'), allow_blank=True)
                ws.add_data_validation(dv_menu)
                dv_menu.add(f"A2:A{max_row+100}")
                
                # Submenu (B) -> Submenu!B
                dv_sub = DataValidation(type="list", formula1=get_list_formula('Submenu', '$B$2:$B$500'), allow_blank=True)
                ws.add_data_validation(dv_sub)
                dv_sub.add(f"B2:B{max_row+100}")
                
            elif sheet_name == "SubmenuItem":
                # Submenu (A)
                dv_sub = DataValidation(type="list", formula1=get_list_formula('Submenu', '$B$2:$B$500'), allow_blank=True)
                ws.add_data_validation(dv_sub)
                dv_sub.add(f"A2:A{max_row+500}")
                
                # Item (C)
                dv_item = DataValidation(type="list", formula1=get_list_formula('Item', '$B$2:$B$2000'), allow_blank=True)
                ws.add_data_validation(dv_item)
                dv_item.add(f"C2:C{max_row+500}")

        # Save to bytes
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

if __name__ == "__main__":
    print("Testing Template-based Builder...")
    builder = ExcelBuilder()
    
    # Dummy Data
    dummy_data = {
        "categories": [{"number": 1, "name": "Food"}, {"number": 2, "name": "Drinks"}],
        "items": [{"number": 100, "name": "Burger", "price": 10.0, "category": "Food", "modifiers": ["Sides"]}],
        "modifier_groups": [
            {"number": 1000, "name": "Sides", "min": 1, "max": 1, "items": [{"name": "Fries", "price": 0, "number":101}]}
        ],
        "submenus": [{"number": 200, "name": "Entrees", "items": ["Burger"]}]
    }
    builder.add_data(dummy_data)
    
    data = builder.build_excel()
    with open("Aloha_Import_Template_Generated.xlsx", "wb") as f:
        f.write(data)
    print("Success! Created 'Aloha_Import_Template_Generated.xlsx' from template.")
