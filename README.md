# emu_assistant_bot

Bu bot yordamida Telegramdagi mijozlar ma'lumotlarini tartiblash, matn yoki rasmdan ajratish va Excel faylga yozib borish mumkin.

## Imkoniyatlar

- `/start` - bot haqida qisqa tushuntirish
- Matndan mijozlarni ajratish
- Rasm, skrinshot yoki qo'lda yozilgan daftar rasmlaridan ma'lumot ajratish
- Bir xabardagi bir nechta mijozni alohida Excel qatorlariga yozish
- Telefonlarni `998XXXXXXXXX` formatiga keltirish
- Noaniq telefon yoki o'qilishi qiyin joylarni `Tekshirish kerak` ustuniga yozish
- `/excel` - tayyor Excel faylni yuborish
- `/clear` - ro'yxatni tozalash
- `/help` - foydalanish bo'yicha yordam

## Excel ustunlari

- No
- Ism familiya
- Telefon raqami
- Manzil
- Qo'shimcha izoh
- Tekshirish kerak

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
```

Botni ishga tushirish:

```bash
python main.py
```

`data/customers.xlsx` fayli avtomatik yaratiladi.

## Railway yoki Render

1. Loyihani GitHub repositoryga yuklang.
2. Railway yoki Renderda yangi Python service yarating.
3. Environment variables bo'limiga quyidagilarni kiriting:
   - `BOT_TOKEN`
   - `OPENAI_API_KEY`
   - `OPENAI_MODEL` ixtiyoriy
4. Start command:

```bash
python main.py
```

Bot long polling orqali ishlaydi, alohida webhook sozlash shart emas.

## Eslatma

Rasm sifati qanchalik yaxshi bo'lsa, natija shunchalik aniq bo'ladi. Qo'lda yozilgan yoki xira ma'lumotlarda bot ehtiyotkorlik bilan noaniq joylarni `Tekshirish kerak` ustuniga yozadi.
