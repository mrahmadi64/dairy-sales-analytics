
📊 درباره پروژه
این پروژه یک تحلیل جامع از داده‌های فروش محصولات لبنی ارائه می‌دهد. با استفاده از این ابزار می‌توانید روندهای فروش، عملکرد محصولات، و کارایی مزارع را تحلیل کرده و گزارش‌های حرفه‌ای تولید کنید.
💾 درباره دیتاست

نام: Dairy Goods Sales Dataset
منبع: Kaggle
نویسنده: Suraj Verma
تعداد رکوردها: 4,325
لایسنس: CC0 (Public Domain)

⚡️ ویژگی‌های اصلی

تحلیل روند فروش با نمودارهای تعاملی
رتبه‌بندی و ارزیابی عملکرد محصولات
تحلیل کارایی مزارع
مدیریت موجودی و ارزیابی ریسک
تحلیل رفتار مشتریان
گزارش‌گیری خودکار در اکسل

🚀 نصب و راه‌اندازی
bash# کلون کردن مخزن
git clone https://github.com/[username]/dairy-market-analysis
cd dairy-market-analysis

# ایجاد محیط مجازی
python -m venv venv
source venv/bin/activate  # Linux/Mac
# یا
venv\Scripts\activate  # Windows

# نصب وابستگی‌ها
pip install -r requirements.txt
📈 نحوه استفاده
pythonfrom src.analyzer import AdvancedDairyAnalyzer

# ایجاد نمونه از کلاس تحلیلگر
analyzer = AdvancedDairyAnalyzer('data/raw/dairy_dataset.csv')

# تولید گزارش
analyzer.generate_excel_report('reports/excel/analysis_report.xlsx')
📁 ساختار پروژه


dairy-market-analysis/  
├── data/                      # پوشه داده‌ها    
│   ├── raw/                   # داده‌های خام  
│   └── processed/             # داده‌های پردازش شده  
├── src/                       # کدهای اصلی  
├── notebooks/                 # نوت‌بوک‌های تحلیلی  
├── reports/                   # گزارش‌ها  
├── tests/                     # تست‌ها  
└── docs/                      # مستندات  
