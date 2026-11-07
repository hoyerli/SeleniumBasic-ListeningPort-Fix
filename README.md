# SeleniumBasic-ListeningPort-Fix

🇹🇷 [Türkçe için tıklayın](README.tr.md)

Fix for the SeleniumBasic error in Excel VBA that started after a Windows update:

```
Run-time error '21':

TimeoutError
The driver failed to open the listening port 127.0.0.1:54996 within 10s
```

(The port number changes on every run.)

## Why it happens

The error is **not** caused by a Chrome / ChromeDriver version mismatch, a blocked port or antivirus:

- ChromeDriver starts normally and answers on the port (`http://127.0.0.1:<port>/status` returns `"ready": true`).
- But SeleniumBasic does not check the port by connecting to it. It reads the Windows TCP table
  (`GetTcpTable`, see [`EndPointExt.cs`](https://github.com/florentbr/SeleniumBasic/blob/master/Selenium/Internal/EndPointExt.cs))
  and looks for a listening entry. After the Windows update, ChromeDriver's listening socket no longer
  shows up there, so SeleniumBasic waits 10 seconds and gives up, even though the driver is running.

The same error is raised outside Excel too (for example from a `.vbs` script), so it is not an Excel / Office problem.

## There are 2 ways to fix it

| | Method | When to use it |
|---|---|---|
| **1** | **Use the `SeleniumFix.bas` module** (recommended) | Keeps Windows up to date. A small code change. |
| **2** | **Uninstall update KB5127936** | No code change, but leaves the PC without that update. |

---

## Method 1 – `SeleniumFix.bas` module (recommended)

The module starts `chromedriver.exe` itself and connects SeleniumBasic to it with `StartRemotely`.
This skips SeleniumBasic's broken port check completely.

### Step 1 – Make sure ChromeDriver matches Chrome

1. In Chrome, open `chrome://settings/help` and note the version (e.g. `156.0.8078.80`).
2. Download the matching **win64** `chromedriver` from
   [Chrome for Testing](https://googlechromelabs.github.io/chrome-for-testing/).
3. Copy `chromedriver.exe` over
   `C:\Users\<your user>\AppData\Local\SeleniumBasic\chromedriver.exe`.
4. Right-click the file → **Properties** → if there is an **Unblock** checkbox at the bottom, tick it → **OK**.

### Step 2 – Download the module

Download [`SeleniumFix.bas`](SeleniumFix.bas): open the file on GitHub → click the **Download raw file** button (⬇).

> If your browser saves it as `SeleniumFix.bas.txt`, rename it to `SeleniumFix.bas`.

### Step 3 – Import the module into Excel

1. Open your Excel file and press **Alt + F11** (VBA editor).
2. In the **Project** window on the left, click your workbook (`VBAProject (YourFile.xlsm)`).
3. Menu **File → Import File...** (or **Ctrl + M**).
4. Select `SeleniumFix.bas` → **Open**.
5. A new module named **SeleniumFix** appears under **Modules**.
6. Check that **Tools → References** has **Selenium Type Library** ticked (it already is if your code worked before).
7. Save the workbook as **.xlsm**.

### Step 4 – Change your code

Only the line that starts the browser changes. Everything else (`.Get`, `.FindElementBy...`, etc.) stays the same.

**Before:**

```vba
Dim drv As New Selenium.ChromeDriver
drv.AddArgument "--start-maximized"
drv.Start "chrome"              ' <-- this line raises error 21
drv.Get "https://www.google.com"
' ...
drv.Quit
```

**After:**

```vba
Dim drv As New Selenium.ChromeDriver
drv.AddArgument "--start-maximized"   ' AddArgument / SetProfile etc. must come BEFORE StartChrome
StartChrome drv                       ' <-- replaces drv.Start "chrome"
drv.Get "https://www.google.com"
' ...
drv.Quit
StopChromeDriver                      ' <-- new: closes chromedriver.exe left in the background
```

Summary of the changes:

| Old | New |
|---|---|
| `drv.Start "chrome"` (or `drv.Start`, `drv.Start "chrome", "https://..."`) | `StartChrome drv` |
| *(nothing)* | `StopChromeDriver` after `drv.Quit` |

If you used `drv.Start "chrome", "https://site.com"` (with a base URL), write `StartChrome drv` and then `drv.Get "https://site.com"`.

### Closing ChromeDriver left in the background

Because we start `chromedriver.exe` ourselves, `drv.Quit` closes Chrome but **not** the driver.
Call `StopChromeDriver` at the end of your macro, and also in your error handler:

```vba
Sub MyMacro()
    Dim drv As New Selenium.ChromeDriver
    On Error GoTo Cleanup

    StartChrome drv
    drv.Get "https://www.google.com"
    ' ... your work ...

Cleanup:
    On Error Resume Next
    drv.Quit
    StopChromeDriver
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub
```

`StopChromeDriver` runs this command hidden:

```bash
taskkill /f /im chromedriver.exe
```

You can also run that command yourself in a Command Prompt if drivers pile up in Task Manager.
`StartChrome` also calls `StopChromeDriver` first, so an old driver never blocks port 9515.

> Note: `taskkill` closes **all** `chromedriver.exe` processes, including ones started by other programs.

### Optional: using a different port

The module uses port **9515**. If something else uses that port, change this line at the top of the module:

```vba
Private Const DRIVER_PORT As Long = 9515
```

---

## Method 2 – Uninstall update KB5127936

If the error started right after **KB5127936** was installed, removing that update brings back the old behaviour,
and your code works without changes.

> ⚠️ This removes security fixes. Windows Update will also try to install the update again,
> so you may need to pause updates. Method 1 is the safer long-term option.

### Option A – From Settings

1. **Start → Settings → Windows Update → Update history**.
2. Scroll down and click **Uninstall updates**.
3. Find **KB5127936** in the list → **Uninstall** → confirm.
4. Restart the computer.

### Option B – From the command line

1. Click Start, type `cmd`, right-click **Command Prompt** → **Run as administrator**.
2. Run:

   ```bash
   wusa /uninstall /kb:5127936
   ```

3. Confirm and restart when asked.

If `wusa` says the update cannot be removed, find the package name and remove it with DISM:

```bash
dism /online /get-packages /format:table | findstr 5127936
```

```bash
dism /online /remove-package /packagename:<package name from the previous command>
```

### Stop it from coming back

**Settings → Windows Update → Pause updates** (choose how many weeks).
When you resume updates, the error will return unless you also apply Method 1.

---

## Tested with

- Windows 11 Pro (build 26300), Office x64
- SeleniumBasic 2.0.9.0
- Chrome / ChromeDriver 156.0.8078.80

The `StartRemotely` approach used by the module was verified on this setup (Chrome opened and loaded a page).

## License

[MIT](LICENSE)
