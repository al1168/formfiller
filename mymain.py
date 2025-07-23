# Example usage
import pandas as pd
from pprint import pprint as pp
from docx import Document
import re

class WordDocumentFiller:
    def __init__(self):
        self.data = pd.DataFrame
        self.template = None
    
    def load_from_excel(self, file_path):
        print(f"Loading data from {file_path}...")
        self.data = pd.read_excel(file_path)
        print(f"Data loaded from {file_path}")
        
    def load_template(self, path):
        self.template = Document(path)
        print(f"Template loaded from {path}")
    def getRowByCenterId(self, center_id):
        """Get a row of data by center ID."""
        row = self.data[self.data['Center_ID'] == center_id]
        if not row.empty:
            return row.iloc[0]
        else:
            print(f"No data found for Center_ID: {center_id}")
            return None
        
    def fill_template(self, center_id, input_date=""):
        dataRow  = self.getRowByCenterId(center_id)
        if dataRow is None: return None

        def safe_str(value):
            if pd.isna(value) or value is None:
                return ""
            return str(value)
        pattern = r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}'
        emergency_text = safe_str(dataRow["Emergency"])
        match = re.search(pattern, emergency_text)
        phone_number = match.group() if match else ""
        placeholder_map = {
            "{FULL_NAME}": safe_str(dataRow["First_Name"]) + " " + safe_str(dataRow["Last_Name"]),
            "{CHINESE_NAME}": safe_str(dataRow["Chinese_Name"]),
            "{DOB}": safe_str(dataRow["DOB"]),
            "{ADDRESS}": safe_str(dataRow["Address"]),
            "{LANGUAGE}": safe_str(dataRow["Language"]),
            "{MEDICAID_ID}": safe_str(dataRow["Medicaid"]),
            "{GENDER}": safe_str(dataRow["Gender"]),
            "{PCP}": safe_str(dataRow["PCP"]),
            "{NAME}": safe_str(dataRow["Emergency"]),  # Emergency contact name
            "{CURRENT_DATE}": input_date if input_date else datetime.now().strftime("%m/%d/%Y"),
            "{COMPANY_ID}": safe_str(dataRow["Member_ID"]),
            "{COMPANY}": safe_str(dataRow["Health_Plan"]),
            "{PHONE}": safe_str(dataRow["Home_Tel"])  if safe_str(dataRow["Home_Tel"]) != "" else safe_str(dataRow["Cell"]),
            "{MEDICARE_ID}": safe_str(dataRow["Medicare"]),
            "{EMERGENCY_PHONE}":  phone_number  # Use .get() for optional fields
        }
        # for i, paragraph in enumerate(self.template.paragraphs):
        #     if paragraph.text.strip():
        #         print(f"Paragraph {i}: '{paragraph.text}'")
        ## iterate through template cells

        for table_idx, table in enumerate(self.template.tables):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    
                    for paragraph in cell.paragraphs:
                        full_text = paragraph.text
                        has_placeholder = any(placeholder in full_text for placeholder in placeholder_map.keys())
                        if has_placeholder:
                            new_text = full_text
                            for placeholder, value in placeholder_map.items():
                                # print(value, placeholder)
                                if placeholder in new_text:
                                    new_text = new_text.replace(placeholder, value.strip())
                                    print(f"✅ Replaced {placeholder} with '{value}' in: '{full_text[:50]}...'")
                                    # replacements_made += 1

    def save_filled_document(self, output_path):
        """Save the filled document to a file."""
        if self.template:
            self.template.save(output_path)
            print(f"Document saved to: {output_path}")
        else:
            print("No template loaded to save.")

def main():
    filler = WordDocumentFiller()
    filler.load_from_excel("ScriptContacts.xlsx")
    filler.load_template('templateCopy.docx')
    filler.fill_template(100,"12-12-1222")

if __name__ == "__main__":
    main()