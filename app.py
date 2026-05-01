from flask import Flask, jsonify, render_template
import pandas as pd
import os

app = Flask(__name__)

def get_excel_value():
    file_path = 'data_live.xlsx'
    try:
        if not os.path.exists(file_path):
            return None, "ملف الإكسيل (data_live.xlsx) غير موجود"

        # قراءة الملف بدون عناوين لضمان دقة أرقام الصفوف
        df = pd.read_excel(file_path, header=None, engine='openpyxl')
        
        # التأكد من وجود 3 صفوف على الأقل
        if len(df) <= 2:
            return None, "الملف لا يحتوي على 3 صفوف"

        # الصف الثالث (Index 2)
        row_3 = df.iloc[2]
        
        val = None
        # العمود F (Index 5)
        if len(row_3) > 5:
            val = row_3.iloc[5]

        # التحقق إذا كانت الخلية فارغة أو غير صالحة
        if pd.isna(val) or str(val).strip() == '':
            # البحث عن أول رقم في الصف
            for item in row_3:
                if pd.notna(item) and isinstance(item, (int, float)):
                    return float(item), None
            return None, "الخلية فارغة ولا يوجد أي رقم في الصف"

        return float(val), None

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

    # منع الكاش نهائياً من طرف السيرفر
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

if __name__ == '__main__':
    # إعداد الـ Host ليكون 0.0.0.0
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)