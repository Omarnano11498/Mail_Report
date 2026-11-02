# نظام أتمتة متابعة بريد Zoho Mail

يجلب الرسائل من Zoho Mail، ويصنّفها، ويرتّبها حسب المرسل، ويصدّر تقريرًا عربيًا إلى ملف Word (.docx).

## 1) التجهيز (مرة واحدة)

1. ثبّت Python 3.9 أو أحدث، ثم المكتبة:
   ```
   pip install python-docx
   ```
2. فعّل الوصول عبر IMAP في Zoho Mail:
   الإعدادات (Settings) ← حسابات البريد (Mail Accounts) ← IMAP ← فعّل IMAP Access.
3. أنشئ كلمة مرور خاصة بالتطبيق (App Password)، وهي إلزامية إذا كان التحقق الثنائي مفعّلًا:
   accounts.zoho.com ← الأمان (Security) ← كلمات مرور التطبيقات (App Passwords).
4. عدّل `config.json`:
   - `email`: بريدك.
   - `imap_host`: الخادم المناسب لمنطقة حسابك:
     | المنطقة | الخادم |
     |---|---|
     | عام/أمريكا | `imap.zoho.com` |
     | أوروبا | `imap.zoho.eu` |
     | الهند | `imap.zoho.in` |
     | حسابات المؤسسات (نطاق خاص) | `imappro.zoho.com` (أو .eu / .in) |
   - `folders`: المجلدات المطلوب متابعتها، مثل `["INBOX", "Sent"]`.

> **الأمان:** لا تكتب كلمة المرور داخل الملف. ضعها في متغير بيئة:
> - Windows: `setx ZOHO_PASSWORD "كلمة-مرور-التطبيق"`
> - macOS/Linux: `export ZOHO_PASSWORD="كلمة-مرور-التطبيق"`

## 2) التشغيل

| الأمر | الوظيفة |
|---|---|
| `python zoho_mail_report.py --demo` | تجربة ببيانات وهمية دون اتصال |
| `python zoho_mail_report.py` | آخر 7 أيام (القيمة في `days_back`) |
| `python zoho_mail_report.py --days 1` | آخر يوم |
| `python zoho_mail_report.py --unseen` | غير المقروءة فقط |
| `python zoho_mail_report.py --new-only` | ما وصل منذ آخر تشغيل فقط (للمتابعة الدورية) |

تُحفظ التقارير في مجلد `reports/` باسم يتضمن التاريخ والوقت.
السكربت يقرأ فقط: لا يحذف، ولا ينقل، ولا يغيّر حالة "مقروءة" لأي رسالة.

## 3) التشغيل التلقائي (المتابعة الدورية)

**Windows — جدولة المهام (Task Scheduler):**
```
schtasks /create /tn "ZohoMailReport" /sc daily /st 08:00 /tr "python C:\zoho\zoho_mail_report.py --new-only --config C:\zoho\config.json"
```

**macOS/Linux — cron** (كل يوم الساعة 8 صباحًا):
```
0 8 * * * cd /path/to/zoho && /usr/bin/python3 zoho_mail_report.py --new-only
```

## 4) تخصيص التصنيف

في `config.json` ← `rules.categories`، لكل تصنيف:
- `sender_patterns`: أجزاء من بريد المرسل (وزنها 3 نقاط)، مثل `"@client-co.com"`.
- `keywords`: كلمات في العنوان أو النص (وزنها نقطة واحدة).

يأخذ كل بريد التصنيف الأعلى نقاطًا، وما لا يطابق شيئًا يُصنَّف "أخرى".
`urgent_keywords` تضع علامة "عاجل" وتجمع هذه الرسائل في قسم مستقل أعلى التقرير.

## محتوى التقرير

1. ملخص تنفيذي: الإجمالي، وعدد المرسلين، وغير المقروءة، والعاجلة.
2. التوزيع حسب التصنيف مع النسب.
3. قائمة الرسائل العاجلة.
4. ترتيب المرسلين حسب حجم المراسلة.
5. تفاصيل كل مرسل: التاريخ، والموضوع، والتصنيف، والحالة، والمرفقات، ومقتطف من النص.
