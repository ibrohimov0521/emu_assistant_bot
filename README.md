# emu_assistant_bot

Bu bot yordamida Telegramdagi mijozlar ma'lumotlarini tartiblash, matn yoki rasmdan ajratish va Excel faylga yozib borish mumkin.

## Imkoniyatlar

- `/start` - bot haqida qisqa tushuntirish
- `/setup` - jo'natuvchi ma'lumotlari va jo'natma parametrlarini qayta sozlash
- `/ofislar` - EMU ofislarini viloyat bo'yicha ko'rish
- `/kalkulyator` - EMU API orqali yetkazib berish narxini hisoblash
- `/ai` - EMU bo'yicha savollar uchun AI yordamchi
- Matndan mijozlarni ajratish
- Rasm, skrinshot yoki qo'lda yozilgan daftar rasmlaridan ma'lumot ajratish
- Bir xabardagi bir nechta mijozni alohida Excel qatorlariga yozish
- Telefonlarni `998XXXXXXXXX` formatiga keltirish
- Jo'natmalarga qaytarilmaydigan shifr berish: `ABC1`, `ABC2`, `ABC3`...
- P ustunga faqat import dasturidagi справочникka mos ruscha shahar/tuman nomini yozish
- `/excel` - tayyor Excel faylni yuborish
- `/shablon` - Excel shablon faylini yuborish
- `/clear` - ro'yxatni tozalash
- `/help` - foydalanish bo'yicha yordam

## Asosiy menyu

Bot ruxsat berilgan foydalanuvchiga quyidagi bo'limlarni ko'rsatadi:

- `Excel ga yig'ish` - yuridik yoki ФИЗ ЛИЦО yo'nalishi bo'yicha jo'natmalarni Excelga yig'ish
- `Ofislar ro'yxati` - oldin viloyat tanlanadi, keyin shu viloyatdagi EMU ofislari chiqadi
- `Kalkulyator` - jo'natuvchi va oluvchi hudud, yetkazish turi va og'irlik bo'yicha narx chiqaradi
- `AI yordamchi` - ofis, filial, shahar/tuman va kalkulyator bo'yicha savollarga yordam beradi
- `Arxiv` - Excel va shablon fayllar
- `Sozlamalar` - jo'natuvchi sozlamalari va ruxsat holati

## Excel ustunlari

- A: Номер
- B: Компания-получатель
- C: ФИО получателя
- D: Адрес получателя
- E: Телефон получателя
- F: Шифр клиента
- G: Масса посылки
- H: Поручение
- I: Количество мест
- J: Штрихкод (№ накладной)
- K: Компания-отправитель
- L: ФИО отправителя
- M: Адрес отправителя
- N: Телефон отправителя
- O: Город-отправитель
- P: Город-получатель
- Q: Оплата получателем

## Ishlash tartibi

Bot Excelga yozishdan oldin quyidagilarni so'raydi:

- jo'natuvchi ism familiyasi
- jo'natuvchi telefon raqami
- jo'natuvchi to'liq manzili
- jo'natuvchi qaysi tuman/shahardanligi rus tilida
- shifr prefixi, masalan `ABC`
- yetkazib berish turi: `ДО ОФИСА` yoki `НА ДОМ`
- `Оплата получателем`: `✅ qo'yilsin` yoki `⬜ qo'yilmasin`
- hamma jo'natmalar uchun bir xil og'irlik
- bir mijozga nechta jo'natma bo'lishi

`ДО ОФИСА` / `НА ДОМ` va `Оплата получателем` qiymatlari Telegram tugmalari orqali tanlanadi.

Shundan keyin mijozlar matn yoki rasm qilib yuboriladi. Bot oluvchilarni ajratib, `templates/yangi_shablon.xlsx` asosida `data/customers.xlsx` faylini yaratadi.

`templates/branch_codes.xlsx` faylidagi `Внутренний код` qiymatlari D ustun uchun ishlatiladi. P ustun o'z nomi bilan qoladi.

`ДО ОФИСА` tanlansa D ustunga aniqlangan viloyat markazining kodi yoziladi. `НА ДОМ` tanlansa D ustunga mijoz yuborgan manzil/oriyentir yoziladi; manzil bo'lmasa aniqlangan tuman nomiga `markazi` qo'shiladi.

## O'rnatish

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

`.env.example` faylidan `.env` yarating:

```env
BOT_TOKEN=telegram_bot_token
OPENAI_API_KEY=openai_api_key
OPENAI_MODEL=gpt-4o-mini
BATCH_CONCURRENCY=10
ADMIN_IDS=6388458077
```

Botni ishga tushirish:

```bash
python main.py
```

`data/customers.xlsx` fayli avtomatik yaratiladi.

## Railway yoki Render

1. Loyihani GitHub repositoryga yuklang.
2. Railway yoki Renderda yangi Python service yarating.
3. Railway/Render ichida `Variables` yoki `Environment variables` bo'limiga quyidagilarni kiriting:
   - `BOT_TOKEN`
   - `OPENAI_API_KEY`
   - `OPENAI_MODEL` ixtiyoriy
   - `BATCH_CONCURRENCY` ixtiyoriy, bir vaqtda nechta xabar tahlil qilinishini belgilaydi
   - `ADMIN_IDS`
4. Start command:

```bash
python main.py
```

Bot long polling orqali ishlaydi, alohida webhook sozlash shart emas.

Railway logida `BOT_TOKEN environment variable sozlanmagan` chiqsa, demak token hali Railway service variables ichiga qo'yilmagan. `.env` fayl GitHubga yuklanmaydi va Railway uni avtomatik ko'rmaydi.

## Admin ruxsati

`ADMIN_IDS` ichiga admin Telegram user ID yoziladi. Bir nechta admin bo'lsa vergul bilan ajrating:

```env
ADMIN_IDS=6388458077
```

Yangi foydalanuvchi `/start` bosganda adminlarga ruxsat so'rovi boradi. Admin tasdiqlagan userlar `data/approved_users.json` faylida saqlanadi.

P ustun (`Город-получатель`) uchun bot `Ферганская область, Учкуприкский район` kabi uzun format yozmaydi. U faqat ichki ro'yxatdagi справочник nomlaridan birini yozadi, masalan: `Учкуприк`, `Бука`, `Тайлак`, `Каракуль`, `Алмазар`.

## Eslatma

Rasm sifati qanchalik yaxshi bo'lsa, natija shunchalik aniq bo'ladi. Qo'lda yozilgan yoki xira ma'lumotlarda bot ehtiyotkorlik bilan noaniq joylarni `Tekshirish kerak` ustuniga yozadi.
