
"""
برنامج البحث عن موظفين في ملف Excel
المبرمج: مساعد الذكاء الصناعي
تاريخ الإنشاء: 2025-04-20
"""

import pandas as pd

# ---------------------------
# 1. قراءة ملف Excel وجمع البيانات
# ---------------------------
def load_excel_data(file_path: str) -> pd.DataFrame:
    """
    دالة لقراءة جميع الجداول من ملف Excel ودمجها في إطار بيانات واحد.
    """
    sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
    all_data = pd.DataFrame()
    
    for sheet_name, df in sheets.items():
        required_columns = ["NAME (ENG)", "ID#", "NATIONALITY", "COMPANY", "POSITION"]
        if all(col in df.columns for col in required_columns):
            all_data = pd.concat([all_data, df], ignore_index=True)
    
    return all_data

# ---------------------------
# 2. دالة البحث عن الموظف
# ---------------------------
def search_employee(data: pd.DataFrame, search_term: str) -> pd.DataFrame:
    """
    دالة للبحث عن الموظف بالاسم أو رقم الهوية.
    """
    mask = (
        data["NAME (ENG)"].astype(str).str.contains(search_term, case=False, na=False) |
        data["ID#"].astype(str).str.contains(search_term, na=False)
    )
    return data[mask]

# ---------------------------
# 3. التشغيل الرئيسي للبرنامج
# ---------------------------
def main():
    file_path = "DUTY ROSTER MAR 2025.V.2.xlsx"  # مسار ملف Excel

    try:
        employee_data = load_excel_data(file_path)
    except FileNotFoundError:
        print("❌ ملف Excel غير موجود. تأكد من المسار.")
        return

    while True:
        search_input = input("\nأدخل اسم الموظف أو رقم الهوية (أو 'خروج' للإنهاء): ").strip()
        if search_input.lower() in ['خروج', 'exit']:
            break
        if not search_input:
            print("⚠️ الرجاء إدخال اسم أو رقم.")
            continue

        results = search_employee(employee_data, search_input)

        if not results.empty:
            print("\n" + "=" * 50)
            print(f"عدد النتائج: {len(results)}")
            print("=" * 50)
            columns_to_display = [
                "NAME (ENG)", "NAME (AR)", "ID#", 
                "NATIONALITY", "COMPANY", "POSITION", "LOCATION"
            ]
            display = results[columns_to_display] if all(col in results.columns for col in columns_to_display) else results
            print(display.to_string(index=False))
        else:
            print("⚠️ لا توجد نتائج.")

if __name__ == "__main__":
    main()
