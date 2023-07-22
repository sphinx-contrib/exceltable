# Test pandas
import pandas as pd

# Test what type of files can be opened with openpyxl
def test_pandas():
    # Try opening:
    # * Excel 97-2003 Workbook (.xls)
    # * Excel Workbook (.xlsx)
    # * Excel Macro-Enabled Workbook (.xlsm)
    # * Excel Workbook Template (.xltx)
    # * Excel Macro-Enabled Workbook Template (.xltm)
    # * Excel Binary Workbook (.xlsb)
    # * OpenDocument Spreadsheet (.ods)
    # * OpenDocument Text (.odt)
    # * OpenDocument Formula (.odf)

    files = ['example/cartoons.xls', 
                'example/cartoons.xlsx', 
                'example/cartoons.xlsm', 
                'example/cartoons.xltx', 
                'example/cartoons.xltm',
                'example/cartoons.xlsb',
                'example/cartoons.ods',
                'example/cartoons.odt',
                'example/cartoons.odf',
    ]
    
    for file in files:
        try:
            df = pd.read_excel(file, sheet_name='quad', usecols='B:C', skiprows=1, header=None)
            print(f'{file} opened successfully!')
            print(df.head())
        except Exception as e:
            print(e)
            print(f'{file} failed to open!\n\n\n')
            pass

if __name__ == '__main__':
    test_pandas()