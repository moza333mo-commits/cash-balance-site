from flask import Flask, jsonify, render_template
import pandas as pd
import os

app = Flask(__name__)

def get_excel_value():
    file_path = 'data.xlsm'  # ✅ تم التعديل هنا

    try:
        if not os.path.exists(file_path):
            return None, "ملف الإكسيل غير موجود"

        df = pd.read_excel(file_path, header=None, engine='openpyxl')

        if len(df) <= 2:
            return None, "الملف لا يحتوي على 3 صفوف"

        row_3 = df.iloc[2]

        val = None

        # العمود F (index 5)
        if len(row_3) > 5:
            val = row_3.iloc[5]

        # 🔥 معالجة القيم الغير صالحة
        if pd.isna(val) or str(val).strip().lower() in ['', '#n/a', 'nan']:
            # ابحث عن أول رقم في الصف
            for item in row_3:
                if pd.notna(item) and isinstance(item, (int, float)):
                    return float(item), None
            return None, "لا يوجد رقم صالح في الصف"

        # 🔥 تحويل آمن للرقم
        try:
            return float(val), None
        except:
            return None, f"القيمة غير رقمية: {val}"

    except Exception as e:
        return None, f"خطأ في قراءة الملف: {str(e)}"


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/data')
def data():
    val, error = get_excel_value()

    if error:
        response = jsonify({"error": error})
    else:
        response = jsonify({"value": val})

    # منع الكاش
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"

    return response


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
