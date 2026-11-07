# SeleniumBasic-ListeningPort-Fix

🇬🇧 [English version](README.md)

Windows güncellemesinden sonra Excel VBA'da SeleniumBasic ile Chrome açılırken çıkan şu hatanın çözümü:

```
Run-time error '21':

TimeoutError
The driver failed to open the listening port 127.0.0.1:54996 within 10s
```

(Port numarası her çalıştırmada değişir.)

## Neden oluyor?

Hatanın sebebi Chrome / ChromeDriver sürüm uyumsuzluğu, kapalı port ya da antivirüs **değildir**:

- ChromeDriver normal şekilde açılıyor ve porttan yanıt veriyor (`http://127.0.0.1:<port>/status` adresi `"ready": true` döndürüyor).
- Ama SeleniumBasic portu, porta bağlanarak kontrol etmiyor. Windows'un TCP tablosunu okuyup
  (`GetTcpTable`, bkz. [`EndPointExt.cs`](https://github.com/florentbr/SeleniumBasic/blob/master/Selenium/Internal/EndPointExt.cs))
  orada dinleyen bir kayıt arıyor. Windows güncellemesinden sonra ChromeDriver'ın dinlediği port bu tabloda
  görünmüyor. SeleniumBasic de sürücü çalıştığı halde 10 saniye bekleyip hata veriyor.

Aynı hata Excel dışında da (örneğin bir `.vbs` dosyasından) çıkıyor; yani sorun Excel / Office'te değil.

## Çözümün 2 yolu var

| | Yöntem | Ne zaman tercih edilir |
|---|---|---|
| **1** | **`SeleniumFix.bas` modülünü kullanmak** (önerilen) | Windows güncel kalır. Kodda küçük bir değişiklik gerekir. |
| **2** | **KB5127936 güncellemesini kaldırmak** | Kod değişmez, ama bilgisayar o güncellemeden mahrum kalır. |

---

## 1. Yol – `SeleniumFix.bas` modülü (önerilen)

Modül `chromedriver.exe`'yi kendisi başlatıyor ve SeleniumBasic'i ona `StartRemotely` ile bağlıyor.
Böylece SeleniumBasic'in bozulan port kontrolü tamamen atlanıyor.

### Adım 1 – ChromeDriver'ın Chrome ile aynı sürümde olduğundan emin olun

1. Chrome'da `chrome://settings/help` adresini açın ve sürümü not edin (ör. `156.0.8078.80`).
2. Aynı sürümün **win64** `chromedriver` dosyasını
   [Chrome for Testing](https://googlechromelabs.github.io/chrome-for-testing/) sayfasından indirin.
3. `chromedriver.exe` dosyasını şu dosyanın üzerine kopyalayın:
   `C:\Users\<kullanıcı adınız>\AppData\Local\SeleniumBasic\chromedriver.exe`
4. Dosyaya sağ tıklayın → **Özellikler** → altta **Engellemeyi kaldır** kutusu varsa işaretleyin → **Tamam**.

### Adım 2 – Modülü indirin

[`SeleniumFix.bas`](SeleniumFix.bas) dosyasını açın → sağ üstteki **Download raw file** (⬇) düğmesine basın.

> Tarayıcı dosyayı `SeleniumFix.bas.txt` olarak kaydederse adını `SeleniumFix.bas` olarak değiştirin.

### Adım 3 – Modülü Excel'e aktarın (Import)

1. Excel dosyanızı açın ve **Alt + F11** tuşlarına basın (VBA düzenleyicisi açılır).
2. Soldaki **Project** penceresinde kendi dosyanıza tıklayın (`VBAProject (DosyaAdınız.xlsm)`).
3. Menüden **File → Import File...** (Türkçe Office'te **Dosya → Dosya Al...**) seçin ya da **Ctrl + M** tuşlarına basın.
4. `SeleniumFix.bas` dosyasını seçin → **Aç**.
5. **Modules** altında **SeleniumFix** adında yeni bir modül görünür.
6. **Tools → References** (**Araçlar → Başvurular**) içinde **Selenium Type Library** işaretli olsun (kodunuz daha önce çalışıyorsa zaten işaretlidir).
7. Dosyayı **.xlsm** (makro içerebilen) olarak kaydedin.

### Adım 4 – Kodunuzu değiştirin

Sadece tarayıcıyı başlatan satır değişiyor. Geri kalan her şey (`.Get`, `.FindElementBy...` vb.) aynı kalıyor.

**Eski kod:**

```vba
Dim drv As New Selenium.ChromeDriver
drv.AddArgument "--start-maximized"
drv.Start "chrome"              ' <-- 21 numaralı hatayı veren satır
drv.Get "https://www.google.com"
' ...
drv.Quit
```

**Yeni kod:**

```vba
Dim drv As New Selenium.ChromeDriver
drv.AddArgument "--start-maximized"   ' AddArgument / SetProfile vb. StartChrome'dan ÖNCE yazılmalı
StartChrome drv                       ' <-- drv.Start "chrome" yerine
drv.Get "https://www.google.com"
' ...
drv.Quit
StopChromeDriver                      ' <-- yeni: arka planda kalan chromedriver.exe'yi kapatır
```

Değişikliklerin özeti:

| Eski | Yeni |
|---|---|
| `drv.Start "chrome"` (ya da `drv.Start`, `drv.Start "chrome", "https://..."`) | `StartChrome drv` |
| *(yok)* | `drv.Quit` satırından sonra `StopChromeDriver` |

`drv.Start "chrome", "https://site.com"` gibi adresli kullanıyorsanız önce `StartChrome drv`, ardından `drv.Get "https://site.com"` yazın.

### Arka planda açık kalan ChromeDriver'ı kapatmak

`chromedriver.exe`'yi artık kendimiz başlattığımız için `drv.Quit` Chrome'u kapatır ama **sürücüyü kapatmaz**.
Makronun sonunda, hata yakalama kısmında da `StopChromeDriver` çağırın:

```vba
Sub Makrom()
    Dim drv As New Selenium.ChromeDriver
    On Error GoTo Temizlik

    StartChrome drv
    drv.Get "https://www.google.com"
    ' ... işlemleriniz ...

Temizlik:
    On Error Resume Next
    drv.Quit
    StopChromeDriver
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub
```

`StopChromeDriver` arka planda gizli olarak şu komutu çalıştırır:

```bash
taskkill /f /im chromedriver.exe
```

Görev Yöneticisi'nde çok sayıda `chromedriver.exe` birikmişse bu komutu Komut İstemi'nde kendiniz de çalıştırabilirsiniz.
`StartChrome` da başlamadan önce `StopChromeDriver` çağırır; böylece eski bir sürücü 9515 portunu meşgul etmez.

> Not: `taskkill` başka programların açtıkları dahil **tüm** `chromedriver.exe` işlemlerini kapatır.

### İsteğe bağlı: farklı port kullanmak

Modül **9515** portunu kullanır. Bu port başka bir program tarafından kullanılıyorsa modülün başındaki şu satırı değiştirin:

```vba
Private Const DRIVER_PORT As Long = 9515
```

---

## 2. Yol – KB5127936 güncellemesini kaldırmak

Hata **KB5127936** yüklendikten hemen sonra başladıysa, bu güncellemeyi kaldırmak eski davranışı geri getirir
ve kodunuz değişiklik yapmadan çalışır.

> ⚠️ Bu işlem güvenlik düzeltmelerini de kaldırır. Windows Update güncellemeyi tekrar yüklemeye çalışacağı için
> güncellemeleri duraklatmanız gerekebilir. Uzun vadede 1. yol daha güvenlidir.

### Seçenek A – Ayarlar üzerinden

1. **Başlat → Ayarlar → Windows Update → Güncelleştirme geçmişi**.
2. Aşağı kaydırıp **Güncelleştirmeleri kaldır**'a tıklayın.
3. Listede **KB5127936**'yı bulun → **Kaldır** → onaylayın.
4. Bilgisayarı yeniden başlatın.

### Seçenek B – Komut satırından

1. Başlat'a `cmd` yazın, **Komut İstemi**'ne sağ tıklayın → **Yönetici olarak çalıştır**.
2. Şu komutu çalıştırın:

   ```bash
   wusa /uninstall /kb:5127936
   ```

3. Onaylayın ve istendiğinde yeniden başlatın.

`wusa` güncellemenin kaldırılamayacağını söylerse paket adını bulup DISM ile kaldırın:

```bash
dism /online /get-packages /format:table | findstr 5127936
```

```bash
dism /online /remove-package /packagename:<önceki komutun verdiği paket adı>
```

### Tekrar yüklenmesini engellemek

**Ayarlar → Windows Update → Güncelleştirmeleri duraklat** (kaç hafta olacağını seçin).
Güncellemeleri yeniden açtığınızda, 1. yolu da uygulamadıysanız hata geri gelir.

---

## Test edilen ortam

- Windows 11 Pro (build 26300), Office x64
- SeleniumBasic 2.0.9.0
- Chrome / ChromeDriver 156.0.8078.80

Modülün kullandığı `StartRemotely` yöntemi bu ortamda doğrulandı (Chrome açıldı ve sayfa yüklendi).

## Lisans

[MIT](LICENSE)
