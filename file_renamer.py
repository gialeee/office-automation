from openpyxl import load_workbook
import os


file_dict = {}

def map_file2key(excel_path):
  wb = load_workbook(excel_path)
  sheet = wb.active
  return file_dict


def rename_file(dir_path, excel_path):
  original = os.path.abspath("")
  new = os.path.abspath("")
  
  os.rename(original, new)
  

if __name__ == "__main__":
  pass
  
