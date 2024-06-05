import os

def find_excel():
        current_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(current_dir, "Contacts.xlsx")
        print(file_path)
        return str(file_path)

find_excel()