Attribute VB_Name = "SeleniumFix"
Option Explicit

' SeleniumBasic-ListeningPort-Fix
' https://github.com/hoyerli/SeleniumBasic-ListeningPort-Fix
'
' Workaround for: "TimeoutError - The driver failed to open the listening port 127.0.0.1:xxxxx within 10s"
' Instead of drv.Start, chromedriver.exe is started here and SeleniumBasic connects with StartRemotely,
' which skips SeleniumBasic's port check (Windows TCP table) that no longer works.
'
' Usage / Kullanim:
'   Dim drv As New Selenium.ChromeDriver
'   ' drv.AddArgument "--start-maximized"   ' optional, BEFORE StartChrome / StartChrome'dan ONCE
'   StartChrome drv                         ' instead of / yerine: drv.Start "chrome"
'   drv.Get "https://www.google.com"
'   ...
'   drv.Quit
'   StopChromeDriver                        ' closes chromedriver.exe / arka plandaki surucuyu kapatir

Private Const DRIVER_PORT As Long = 9515

Public Sub StartChrome(drv As Object)
    StopChromeDriver
    Shell """" & Environ$("LOCALAPPDATA") & "\SeleniumBasic\chromedriver.exe"" --port=" & DRIVER_PORT, vbHide

    If Not WaitForDriver(10) Then
        Err.Raise vbObjectError + 513, "StartChrome", "chromedriver did not respond on port " & DRIVER_PORT & "."
    End If

    drv.StartRemotely "http://localhost:" & DRIVER_PORT & "/", "chrome"
End Sub

' Closes every chromedriver.exe left running in the background.
Public Sub StopChromeDriver()
    CreateObject("WScript.Shell").Run "taskkill /f /im chromedriver.exe", 0, True
End Sub

' Waits until http://127.0.0.1:<port>/status answers.
Private Function WaitForDriver(timeoutSec As Long) As Boolean
    Dim http As Object, t As Single, w As Single
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    t = Timer
    Do
        On Error Resume Next
        http.Open "GET", "http://127.0.0.1:" & DRIVER_PORT & "/status", False
        http.setTimeouts 500, 500, 500, 500
        http.send
        If Err.Number = 0 Then
            If http.Status = 200 Then WaitForDriver = True: Exit Function
        End If
        Err.Clear
        On Error GoTo 0

        w = Timer
        Do While Timer - w < 0.25 And Timer >= w
            DoEvents
        Loop
    Loop While Timer - t < timeoutSec And Timer >= t
End Function
