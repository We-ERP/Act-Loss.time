# Call Center Data Structure Automation

نظام معالجة وتحديث تقارير خدمة العملاء أوتوماتيكياً.

## هيكل الملفات والمجلدات:
- `index.html` : واجهة المستخدم بالكامل، ويشغّل الملفات الفعالة `style.css` و`app.js`.
- `style.css` : تصميم وتنسيق الواجهة.
- `app.js` : منطق المعالجة والتصدير المباشر من المتصفح، مع دعم رفع `Schedule / Scheduled Time per Agent`.
- `scripts/process_data.py` : سكربت معالجة البيانات بالسيرفر.
- `data/` : المجلد الذكي لملفات UTL, IR, Compensation, Schedule.
- `templates/ Structure.xlsx` : الشيت الأساسي الهيكلي المراد تحديثه.
- `output/ Final_Report.xlsx` : الناتج المحدث النهائي تلقائياً.

## التشغيل على GitHub Pages:
قم بتفعيل GitHub Pages من إعدادات Repository وقم بفتح رابط الصفحة لاستخدام الواجهة مباشرة عبر المتصفح دون الحاجة لتثبيت أي برامج!

## ملاحظات الاستخدام:
- سيتم استخدام الملف `STR Loss.xlsx` الموجود في جذر المستودع تلقائياً كملف Structure، مع بقاء الرفع اليدوي اختيارياً.
- يمكن رفع ملف `Schedule / Scheduled Time per Agent` بصيغ `XLSX / XLS / CSV`، ويقوم التطبيق باكتشاف صف العناوين تلقائياً حتى لو سبقته صفوف metadata.