# emu_assistant_bot

Bu bot yordamida Telegramdagi mijozlar ma'lumotlarini tartiblash, matn yoki rasmdan ajratish va Excel faylga yozib borish mumkin.

## Imkoniyatlar

- `/start` - bot haqida qisqa tushuntirish
- `/setup` - jo'natuvchi ma'lumotlari va jo'natma parametrlarini qayta sozlash
- Matndan mijozlarni ajratish
- Rasm, skrinshot yoki qo'lda yozilgan daftar rasmlaridan ma'lumot ajratish
- Bir xabardagi bir nechta mijozni alohida Excel qatorlariga yozish
- Telefonlarni `998XXXXXXXXX` formatiga keltirish
- Jo'natmalarga qaytarilmaydigan shifr berish: `ABC1`, `ABC2`, `ABC3`...
- Oluvchi hududini rus tilida AI orqali aniqlash
- `/excel` - tayyor Excel faylni yuborish
- `/clear` - ro'yxatni tozalash
- `/help` - foydalanish bo'yicha yordam

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
- jo'natuvchi shahri rus tilida
- shifr prefixi, masalan `ABC`
- `Оплата получателем`: `True` yoki `False`
- hamma jo'natmalar uchun bir xil og'irlik
- bir mijozga nechta jo'natma bo'lishi

Shundan keyin mijozlar matn yoki rasm qilib yuboriladi. Bot oluvchilarni ajratib, `templates/yangi_shablon.xlsx` asosida `data/customers.xlsx` faylini yaratadi.

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
TOWNLIST_URL=
TOWNLIST_AUTH_EXTRA=
TOWNLIST_COUNTRY=
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
- `TOWNLIST_URL` ixtiyoriy, `Справочник городов` XML API endpointi
- `TOWNLIST_AUTH_EXTRA` ixtiyoriy, API auth `extra` qiymati
- `TOWNLIST_COUNTRY` ixtiyoriy, справочник country filteri
4. Start command:

```bash
python main.py
```

Bot long polling orqali ishlaydi, alohida webhook sozlash shart emas.

Railway logida `BOT_TOKEN environment variable sozlanmagan` chiqsa, demak token hali Railway service variables ichiga qo'yilmagan. `.env` fayl GitHubga yuklanmaydi va Railway uni avtomatik ko'rmaydi.

P ustun (`Город-получатель`) uchun bot viloyat+tuman matnini emas, `Справочник городов` dagi `town/name` formatiga o'xshash aholi punkti nomini yozadi. Agar `TOWNLIST_URL` sozlansa, bot XML API orqali qidirib, birinchi mos `town/name` qiymatini P ustunga yozadi.

## Eslatma

Rasm sifati qanchalik yaxshi bo'lsa, natija shunchalik aniq bo'ladi. Qo'lda yozilgan yoki xira ma'lumotlarda bot ehtiyotkorlik bilan noaniq joylarni `Tekshirish kerak` ustuniga yozadi.
