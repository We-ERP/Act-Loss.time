# CC Performance Dashboard

تقرير يومي لكل موظف: **Assigning Tkts • TKT • System • Talk Time • Tele-SCH • Comp • Loss Time**

## هيكل المستودع
```
├── index.html                      الواجهة
├── style.css                       التصميم
├── app.js                          المعالجة والتصدير من المتصفح
├── config.json                     إعدادات المعادلة + الأكواد الافتراضية
├── STR Loss.xlsx                   شيت الـ Structure (بيتحدّث هنا في المستودع)
├── scripts/process_data.py         نفس المنطق على السيرفر (GitHub Actions)
├── data/                           UTL / IR / Compensation / Schedule (للسيرفر)
├── output/Final_Report.xlsx        الناتج التلقائي
├── tests/structure-master-behavior.test.js
└── .github/workflows/auto_process.yml
```

## المعادلات
| العمود | الحساب |
|---|---|
| Assigning Tkts | عدد صفوف IR اللي `assigned_to` (عمود X) = TTS User، لكل يوم |
| TKT | عدد صفوف IR اللي `added_by` (عمود Y) = TTS User، لكل يوم |
| التاريخ | `added_on` (عمود Z) بصيغة `M/D/YYYY h:mm:ss AM` |
| System | `TKT × 0.00104166666666667` يوم = `TKT × 90 ثانية` |
| Talk Time | UTL بالـ Login ID: `Hold + Other + AUXOUTOFFTIME + ACWOUTOFFTIME` |
| Tele-SCH | مجموع Duration للأكواد المختارة من Schedule (أو إجمالي اليوم لو ماتحددش أكواد) |
| Comp | Compensation بالـ Teleopti ID: `Comp_Du` |
| **Loss Time** | `Tele-SCH × 90% − (System + Talk Time + Comp)` ولا يقل عن صفر |

غير النسبة أو ثواني التذكرة من `config.json` (`teleSchFactor`, `secondsPerTicket`).

## Schedule (نفس منطق الماكرو)
الملف الخام بيتقرأ بالمواقع: **B** = ID + الاسم (بداية بلوك الموظف) • **C** = التاريخ أو كود النشاط • **J** = Duration.
بعد رفع الملف بتظهر الأكواد كأزرار، اختار اللي تتحسب في Tele-SCH (الاختيار بيتحفظ في المتصفح).
عشان تثبّتها للجميع/للسيرفر حطها في `config.json`:
```json
{ "scheduleCodes": ["Phone", "Chat"] }
```
بيدعم كمان جدول مسطح فيه `Agent / Date / Scheduled time`.

## التشغيل
- **المتصفح:** فعّل GitHub Pages وافتح الرابط. `STR Loss.xlsx` بيتحمّل تلقائيًا، وارفع باقي الشيتات واضغط "بدء المعالجة".
- **السيرفر:** ارفع ملفات `data/` (`UTL.xlsx`, `IR.xlsx`, `Compensation.xlsx`, `Schedule.xlsx`) والـ Action هيحدّث `output/Final_Report.xlsx`.
- **محليًا:** `pip install -r requirements.txt && python scripts/process_data.py`
- **اختبارات:** `node tests/structure-master-behavior.test.js`
