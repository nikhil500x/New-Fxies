'==============================================================
' Fix-MSUpdate.vbs
'
' Remediation for the "E-statement.vbs" downloader:
'   staging  : C:\Users\Public\Documents\MSUpdate_<5 random digits>\
'   dropped  : Lo.zip, setup1.vbs
'   source   : https://jb.mywwjj.xyz/sys/D/1/c2b8.zip
'   delivery : cmd /c curl|bitsadmin|powershell|certutil  (hidden)
'   handoff  : wscript.exe setup1.vbs
'
' USAGE (run from an elevated command prompt):
'   cscript //nologo Fix-MSUpdate.vbs          -> report only, changes nothing
'   cscript //nologo Fix-MSUpdate.vbs /fix     -> quarantine + remove
'
' Folders are MOVED to C:\Quarantine_MSUpdate\, never deleted, so the
' sample survives for analysis. Review the report before using /fix.
'==============================================================
Option Explicit

'--- run under cscript so output goes to the console, not popups ---
If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
    Dim relaunch, arg
    relaunch = "cscript.exe //nologo """ & WScript.ScriptFullName & """"
    For Each arg In WScript.Arguments
        relaunch = relaunch & " " & arg
    Next
    CreateObject("WScript.Shell").Run relaunch, 1, False
    WScript.Quit 0
End If

Const QUAR_BASE  = "C:\Quarantine_MSUpdate"
Const BAD_URL    = "https://jb.mywwjj.xyz/sys/D/1/c2b8.zip"
Const BITS_JOB   = "UpdateJob"
Const HKCU       = &H80000001
Const HKLM       = &H80000002

Dim IOC
IOC = Array("msupdate_", "setup1.vbs", "lo.zip", "mywwjj")

Dim fso, sh, reg, HaveReg, DoFix, gFound, gActed, LogFile, Stamp

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

' VBScript does not short-circuit Or, so guard the provider with a flag
' rather than testing the object variable later.
HaveReg = False
On Error Resume Next
Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
If Err.Number = 0 Then HaveReg = True
Err.Clear
On Error GoTo 0

gFound = 0
gActed = 0
Stamp  = MakeStamp()
DoFix  = HasArg("/fix")

Set LogFile = fso.OpenTextFile( _
    fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), _
                  "MSUpdate_Cleanup_" & Stamp & ".log"), 2, True)

Say "=============================================================="
Say " MSUpdate downloader remediation   " & Now
If DoFix Then
    Say " MODE: /fix  -- artifacts WILL be quarantined and removed"
Else
    Say " MODE: report only -- nothing will be changed. Use /fix to act."
End If
Say "=============================================================="
Say ""

If Not IsAdmin() Then
    Say "[!] Not running elevated. HKLM keys, machine-wide scheduled tasks"
    Say "    and other users' BITS jobs will be missed. Re-run as Administrator."
    Say ""
End If

StopMaliciousProcesses
SweepStagingFolders
SweepStrayDrops
ClearBitsJob
ClearCertutilCache
SweepRunKeys
SweepStartupFolders
SweepScheduledTasks

Say ""
Say "=============================================================="
Say " Artifacts found : " & gFound
If DoFix Then
    Say " Actions taken   : " & gActed
    If gFound > 0 Then
        Say ""
        Say " Quarantine: " & QUAR_BASE
        Say " Reboot, then re-run in report mode to confirm it is clean."
    End If
Else
    Say " Actions taken   : 0 (report mode)"
    If gFound > 0 Then Say " Re-run with /fix to remediate."
End If
Say "=============================================================="

LogFile.Close
WScript.Quit 0

'==================== steps ====================

'--- 1. kill script hosts running the dropped payload -----------
Sub StopMaliciousProcesses()
    Say "[1] Running script-host processes"
    Dim wmi, procs, p, hit
    hit = 0
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number <> 0 Then
        Say "    - WMI unavailable (" & Err.Description & "), skipped"
        Err.Clear : On Error GoTo 0 : Say "" : Exit Sub
    End If
    Set procs = wmi.ExecQuery( _
        "SELECT ProcessId,Name,CommandLine FROM Win32_Process " & _
        "WHERE Name='wscript.exe' OR Name='cscript.exe' OR Name='bitsadmin.exe'")
    On Error GoTo 0

    For Each p In procs
        If IsIoc(p.CommandLine) Then
            hit = hit + 1
            gFound = gFound + 1
            Say "    ! PID " & p.ProcessId & "  " & Trim(p.CommandLine)
            If DoFix Then
                On Error Resume Next
                p.Terminate
                If Err.Number = 0 Then
                    Say "      -> terminated"
                    gActed = gActed + 1
                Else
                    Say "      -> could not terminate: " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next
    If hit = 0 Then Say "    - none"
    Say ""
End Sub

'--- 2. the staging folders themselves ---------------------------
Sub SweepStagingFolders()
    Say "[2] Staging folders  C:\Users\Public\Documents\MSUpdate_*"
    Dim base, root, f, hit
    hit = 0
    base = sh.ExpandEnvironmentStrings("%PUBLIC%")
    If base = "%PUBLIC%" Or base = "" Then base = "C:\Users\Public"
    base = fso.BuildPath(base, "Documents")

    If Not fso.FolderExists(base) Then
        Say "    - " & base & " does not exist"
        Say "" : Exit Sub
    End If

    Set root = fso.GetFolder(base)
    For Each f In root.SubFolders
        If LCase(Left(f.Name, 9)) = "msupdate_" Then
            hit = hit + 1
            gFound = gFound + 1
            Say "    ! " & f.Path
            ListContents f
            If DoFix Then QuarantineFolder f.Path
        End If
    Next
    If hit = 0 Then Say "    - none"
    Say ""
End Sub

'--- 3. loose copies of the payload elsewhere --------------------
Sub SweepStrayDrops()
    Say "[3] Stray Lo.zip / setup1.vbs in common drop locations"
    Dim places, i, d, f, hit
    hit = 0
    places = Array( _
        sh.ExpandEnvironmentStrings("%TEMP%"), _
        sh.ExpandEnvironmentStrings("%windir%\Temp"), _
        sh.ExpandEnvironmentStrings("%PUBLIC%"), _
        sh.ExpandEnvironmentStrings("%APPDATA%"), _
        sh.ExpandEnvironmentStrings("%LOCALAPPDATA%"))

    For i = 0 To UBound(places)
        If places(i) <> "" And fso.FolderExists(places(i)) Then
            On Error Resume Next
            Set d = fso.GetFolder(places(i))
            For Each f In d.Files
                If LCase(f.Name) = "setup1.vbs" Or LCase(f.Name) = "lo.zip" Then
                    hit = hit + 1
                    gFound = gFound + 1
                    Say "    ! " & f.Path & "  (" & f.Size & " bytes)"
                    If DoFix Then QuarantineFile f.Path
                End If
            Next
            Err.Clear
            On Error GoTo 0
        End If
    Next
    If hit = 0 Then Say "    - none"
    Say ""
End Sub

'--- 4. the BITS transfer job ------------------------------------
Sub ClearBitsJob()
    Say "[4] BITS job """ & BITS_JOB & """"
    Dim out
    out = RunCapture("bitsadmin.exe /list /allusers")
    If InStr(LCase(out), LCase(BITS_JOB)) > 0 Then
        gFound = gFound + 1
        Say "    ! job present"
        If DoFix Then
            RunCapture "bitsadmin.exe /cancel " & BITS_JOB
            Say "      -> cancel issued (jobs owned by other users need their own session)"
            gActed = gActed + 1
        End If
    Else
        Say "    - not present"
    End If
    Say ""
End Sub

'--- 5. certutil URL cache ---------------------------------------
Sub ClearCertutilCache()
    Say "[5] certutil URL cache entry"
    If DoFix Then
        RunCapture "certutil.exe -urlcache " & BAD_URL & " delete"
        Say "    -> delete issued for " & BAD_URL
        gActed = gActed + 1
    Else
        Say "    - would delete cache entry for " & BAD_URL
        Say "      (also check %LOCALAPPDATA%\Microsoft\Windows\INetCache)"
    End If
    Say ""
End Sub

'--- 6. Run / RunOnce persistence --------------------------------
Sub SweepRunKeys()
    Say "[6] Run / RunOnce registry values"
    If Not HaveReg Then
        Say "    - registry provider unavailable, skipped" : Say "" : Exit Sub
    End If

    Dim hives, paths, h, i, hit
    hit = 0
    hives = Array(HKCU, HKLM)
    paths = Array( _
        "Software\Microsoft\Windows\CurrentVersion\Run", _
        "Software\Microsoft\Windows\CurrentVersion\RunOnce", _
        "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run", _
        "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\RunOnce")

    For Each h In hives
        For i = 0 To UBound(paths)
            hit = hit + ScanRunKey(h, paths(i))
        Next
    Next
    If hit = 0 Then Say "    - none"
    Say ""
End Sub

Function ScanRunKey(hive, path)
    Dim names, types, i, val, hit
    hit = 0
    ScanRunKey = 0
    On Error Resume Next
    If reg.EnumValues(hive, path, names, types) <> 0 Then
        Err.Clear : On Error GoTo 0 : Exit Function
    End If
    On Error GoTo 0
    If Not IsArray(names) Then Exit Function

    For i = 0 To UBound(names)
        val = ""
        On Error Resume Next
        reg.GetStringValue hive, path, names(i), val
        On Error GoTo 0
        If IsIoc(val) Or IsIoc(names(i)) Then
            hit = hit + 1
            gFound = gFound + 1
            Say "    ! " & HiveName(hive) & "\" & path
            Say "        " & names(i) & " = " & val
            If DoFix Then
                On Error Resume Next
                If reg.DeleteValue(hive, path, names(i)) = 0 Then
                    Say "        -> deleted"
                    gActed = gActed + 1
                Else
                    Say "        -> delete failed (need elevation?)"
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next
    ScanRunKey = hit
End Function

'--- 7. Startup folders ------------------------------------------
Sub SweepStartupFolders()
    Say "[7] Startup folders"
    Dim dirs, i, d, f, hit, body
    hit = 0
    dirs = Array(sh.SpecialFolders("Startup"), sh.SpecialFolders("AllUsersStartup"))

    For i = 0 To UBound(dirs)
        If dirs(i) <> "" And fso.FolderExists(dirs(i)) Then
            Set d = fso.GetFolder(dirs(i))
            For Each f In d.Files
                body = ""
                If f.Size < 200000 Then body = ReadTextSafe(f.Path)
                If IsIoc(f.Name) Or IsIoc(body) Then
                    hit = hit + 1
                    gFound = gFound + 1
                    Say "    ! " & f.Path
                    If DoFix Then QuarantineFile f.Path
                End If
            Next
        End If
    Next
    If hit = 0 Then Say "    - none"
    Say ""
End Sub

'--- 8. Scheduled tasks ------------------------------------------
Sub SweepScheduledTasks()
    Say "[8] Scheduled tasks"
    Dim svc, hit
    hit = 0
    On Error Resume Next
    Set svc = CreateObject("Schedule.Service")
    svc.Connect
    If Err.Number <> 0 Then
        Say "    - task scheduler unavailable (" & Err.Description & "), skipped"
        Err.Clear : On Error GoTo 0 : Say "" : Exit Sub
    End If
    On Error GoTo 0

    hit = ScanTaskFolder(svc.GetFolder("\"))
    If hit = 0 Then Say "    - none"
    Say ""
End Sub

Function ScanTaskFolder(folder)
    Dim t, sub_, hit
    hit = 0
    ScanTaskFolder = 0
    On Error Resume Next
    For Each t In folder.GetTasks(1)
        If IsIoc(t.Xml) Then
            hit = hit + 1
            gFound = gFound + 1
            Say "    ! " & t.Path
            If DoFix Then
                folder.DeleteTask t.Name, 0
                If Err.Number = 0 Then
                    Say "      -> deleted"
                    gActed = gActed + 1
                Else
                    Say "      -> delete failed (need elevation?)"
                    Err.Clear
                End If
            End If
        End If
    Next
    For Each sub_ In folder.GetFolders(0)
        hit = hit + ScanTaskFolder(sub_)
    Next
    Err.Clear
    On Error GoTo 0
    ScanTaskFolder = hit
End Function

'==================== helpers ====================

Sub ListContents(folder)
    Dim f, sf
    On Error Resume Next
    For Each f In folder.Files
        Say "        - " & f.Name & "  (" & f.Size & " bytes)"
    Next
    For Each sf In folder.SubFolders
        Say "        - " & sf.Name & "\"
    Next
    Err.Clear
    On Error GoTo 0
End Sub

Sub QuarantineFolder(path)
    Dim dest
    EnsureQuarantine
    dest = fso.BuildPath(QUAR_BASE, fso.GetBaseName(path) & "_" & Stamp)
    On Error Resume Next
    ClearAttributes path
    fso.MoveFolder path, dest
    If Err.Number = 0 Then
        Say "      -> quarantined to " & dest
        gActed = gActed + 1
    Else
        Say "      -> move failed (" & Err.Description & "); file may be locked"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub QuarantineFile(path)
    Dim dest
    EnsureQuarantine
    dest = fso.BuildPath(QUAR_BASE, fso.GetFileName(path) & "_" & Stamp)
    On Error Resume Next
    fso.GetFile(path).Attributes = 0
    fso.MoveFile path, dest
    If Err.Number = 0 Then
        Say "      -> quarantined to " & dest
        gActed = gActed + 1
    Else
        Say "      -> move failed (" & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub EnsureQuarantine()
    On Error Resume Next
    If Not fso.FolderExists(QUAR_BASE) Then fso.CreateFolder QUAR_BASE
    Err.Clear
    On Error GoTo 0
End Sub

' hidden / read-only / system attributes would block MoveFolder
Sub ClearAttributes(path)
    Dim f, sf, d
    On Error Resume Next
    Set d = fso.GetFolder(path)
    d.Attributes = 0
    For Each f In d.Files
        f.Attributes = 0
    Next
    For Each sf In d.SubFolders
        ClearAttributes sf.Path
    Next
    Err.Clear
    On Error GoTo 0
End Sub

Function IsIoc(s)
    Dim i, t
    IsIoc = False
    If IsNull(s) Or IsEmpty(s) Then Exit Function
    t = LCase(CStr(s))
    If t = "" Then Exit Function
    For i = 0 To UBound(IOC)
        If InStr(t, IOC(i)) > 0 Then
            IsIoc = True
            Exit Function
        End If
    Next
End Function

Function ReadTextSafe(path)
    Dim ts
    ReadTextSafe = ""
    On Error Resume Next
    Set ts = fso.OpenTextFile(path, 1, False)
    If Err.Number = 0 Then
        ReadTextSafe = ts.ReadAll
        ts.Close
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function RunCapture(cmd)
    Dim exec
    RunCapture = ""
    On Error Resume Next
    Set exec = sh.Exec("cmd.exe /c " & cmd & " 2>&1")
    If Err.Number <> 0 Then
        Err.Clear : On Error GoTo 0 : Exit Function
    End If
    Do While exec.Status = 0
        WScript.Sleep 100
    Loop
    RunCapture = exec.StdOut.ReadAll
    Err.Clear
    On Error GoTo 0
End Function

Function IsAdmin()
    Dim out
    out = RunCapture("net session")
    IsAdmin = (InStr(LCase(out), "access is denied") = 0)
End Function

Function HiveName(h)
    If h = HKCU Then HiveName = "HKCU" Else HiveName = "HKLM"
End Function

Function MakeStamp()
    Dim d
    d = Now
    MakeStamp = Year(d) & Pad(Month(d)) & Pad(Day(d)) & "_" & _
                Pad(Hour(d)) & Pad(Minute(d)) & Pad(Second(d))
End Function

Function Pad(n)
    If n < 10 Then Pad = "0" & n Else Pad = CStr(n)
End Function

Sub Say(msg)
    WScript.Echo msg
    On Error Resume Next
    LogFile.WriteLine msg
    Err.Clear
    On Error GoTo 0
End Sub
