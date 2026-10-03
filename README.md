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
| Tele-SCH | مجموع Duration للأكواد المختارة × 90% (بيظهر بعد الخصم، مثلاً 8:00 ← 7:12) |
| Comp | Compensation بالـ Teleopti ID: `Comp_Du` |
| **Loss Time** | `Tele-SCH − (System + Talk Time + Comp)` ولا يقل عن صفر |

غير النسبة أو ثواني التذكرة من `config.json` (`teleSchFactor`, `secondsPerTicket`).

## Schedule
بيتعرف تلقائيًا على شكلين:
1. **شيت `Final` بعد الماكرو:** أعمدة `ID | Date | Duration | Agent Name | TL | Code`.
2. **التقرير الخام (RD):** **B** = ID + الاسم • **C** = التاريخ أو كود النشاط • **J** = Duration.

الأكواد المعتمدة (في `config.json`): `Covering E.C` • `Covering HSH instability` • `Covering PC Pro` • `Phone`.
بعد رفع الملف بتظهر الأكواد كأزرار تقدر تغيّر الاختيار منها، والاختيار بيتحفظ في المتصفح.

## التصدير
زر Excel بيطلّع نفس شكل الشيت: الأعمدة الثابتة (Teleopti ID, Login ID, Agent Name, TTS User, TL Name, Status)، صف التواريخ المدموج، عناوين بنفسجي (Assigning Tkts برتقالي)، وLoss Time أخضر/أحمر.
كل يوم 7 أعمدة في مجموعة Outline: اليوم المطوي في الصفحة بيظهر منه Loss Time بس، واليوم المفتوح بيظهر كامل.

## تشخيص IR
بعد المعالجة شريط الحالة بيعرض: عدد صفوف IR، كام صف `added_by` / `assigned_to` طابق TTS User، وكام تاريخ اتقرا ومداه. لو مفيش مطابقة بيعرض أمثلة من القيم.

## التشغيل
- **المتصفح:** فعّل GitHub Pages وافتح الرابط. `STR Loss.xlsx` بيتحمّل تلقائيًا، وارفع باقي الشيتات واضغط "بدء المعالجة".
- **السيرفر:** ارفع ملفات `data/` (`UTL.xlsx`, `IR.xlsx`, `Compensation.xlsx`, `Schedule.xlsx`) والـ Action هيحدّث `output/Final_Report.xlsx`.
- **محليًا:** `pip install -r requirements.txt && python scripts/process_data.py`
- **اختبارات:** `node tests/structure-master-behavior.test.js`
