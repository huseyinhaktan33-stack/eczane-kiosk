# Gerçek Build/Host Kurulumu — Adım Adım

Bu klasör artık gerçek bir React (Vite) projesi. Her değişiklikten sonra
benim elle derleyip sana dosya vermeme gerek kalmayacak — GitHub'a
gönderdiğin an Netlify kendisi derleyip yayınlayacak.

## 1. Bilgisayarında hızlı bir test (opsiyonel ama önerilir)

Node.js kuruluysa (yoksa nodejs.org'dan indir), bu klasörde:

```bash
npm install
npm run dev
```

Tarayıcıda `http://localhost:5173` açılır, her şey çalışıyor mu kontrol et.
`npm install` internet gerektirir (paketleri indirir) — benim burada
yapamadığım şey tam olarak bu, sende sorunsuz çalışacaktır.

## 2. GitHub'a yükle

1. https://github.com → sağ üstten **"New repository"**
2. İsim ver (örn. `eczane-kiosk`), **Private** seçebilirsin, "Create repository"
3. Bilgisayarında bu klasörde:

```bash
git init
git add .
git commit -m "İlk sürüm"
git branch -M main
git remote add origin https://github.com/KULLANICI_ADIN/eczane-kiosk.git
git push -u origin main
```

(`KULLANICI_ADIN` kısmını kendi GitHub kullanıcı adınla değiştir. Git
kurulu değilse git-scm.com'dan indir.)

## 3. Netlify'ı GitHub'a bağla

1. https://app.netlify.com → **"Add new site"** → **"Import an existing project"**
2. **GitHub**'ı seç, izin ver, az önce oluşturduğun `eczane-kiosk` reposunu seç
3. Build ayarları otomatik doldurulmalı (bu klasördeki `netlify.toml` sayesinde):
   - Build command: `npm run build`
   - Publish directory: `dist`
4. **"Deploy site"** butonuna bas

## 4. Supabase bilgilerini Netlify'a gir (önemli!)

1. Netlify'da siten → **Site configuration** → **Environment variables**
2. **"Add a variable"** ile ikisini ekle:
   - `VITE_SUPABASE_URL` = `https://wtbhyeshiqufnjagxvxu.supabase.co`
   - `VITE_SUPABASE_KEY` = `sb_publishable_SMPQ0erqqRNEoboVnlvIew_yxsA2uRj`
3. **Deploys** sekmesinden **"Trigger deploy"** → **"Deploy site"** (değişkenlerin
   etkili olması için bir kere yeniden derletmen gerekiyor)

## Bundan sonrası

Artık kodda bir değişiklik yapmak istediğinde (benimle ya da başka biriyle):

```bash
git add .
git commit -m "değişiklik açıklaması"
git push
```

Bu kadar. Netlify otomatik olarak algılayıp birkaç dakikada yeniden
derleyip yayınlayacak. Benim CDN'lere güvenerek tek dosya derlediğim
kırılgan yöntem artık tarihe karıştı.

## Klasördeki dosyalar ne işe yarıyor

- `src/App.jsx` — uygulamanın tüm kodu (benim sana verdiğim `eczane-kiosk-v9.jsx` ile birebir aynı)
- `src/main.jsx` — React'i sayfaya bağlayan küçük giriş dosyası
- `index.html` — Vite'ın giriş sayfası (PWA meta etiketleri burada)
- `public/manifest.json`, `public/sw.js`, `public/icon-*.png` — PWA/"Ana ekrana ekle" dosyaları
- `package.json` — hangi kütüphanelerin kullanıldığı
- `vite.config.js` — build ayarları
- `netlify.toml` — Netlify'a nasıl derleyeceğini söyler
- `.env.example` — Supabase bilgilerinin örneği (gerçek `.env` dosyasını sen oluşturacaksın, Git'e yüklenmez)

## Sık karşılaşılabilecek sorunlar

- **"npm: command not found"** → Node.js kurulu değil, nodejs.org'dan indir.
- **Netlify'da build hatası** → Deploys sekmesinden ilgili deploy'a tıklayıp
  log'u oku; genelde eksik bir environment variable'dır (Adım 4).
- **Site açılıyor ama Supabase'e bağlanamıyor** → Adım 4'teki environment
  variable'ları girip **yeniden deploy tetiklemeyi** unutma, otomatik
  uygulanmaz.
