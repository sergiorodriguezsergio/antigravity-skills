
import openpyxl

def test_excel_connection(excel_path):
    print(f"Intentando conectar con: {excel_path}")
    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True)
        print("\n✅ Conexión exitosa.")
        print(f"Hojas encontradas en el archivo: {wb.sheetnames}")
        wb.close()
    except Exception as e:
        print(f"\n❌ Error al conectar con el archivo: {e}")

if __name__ == "__main__":
    PATH_EXCEL = r'\\172.16.80.129\g\EXCEL PROGRAMAS\PICK AND ROLL\DATOS PICK AND ROLL.xlsx'
    test_excel_connection(PATH_EXCEL)
