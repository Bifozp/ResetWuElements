' Windows Update Elements Reset Script
' ( for error code 0xc1900204 )
'
'  -- 2021/06/01
'
' Reference: https://ugetfix.com/ask/how-to-fix-windows-10-update-error-code-0xc1900204/
'
' License : MIT
'

Option Explicit
Const DebugMode = False
Dim Wmi
Set Wmi = GetObject("winmgmts:\\.\root\CIMV2")

' UAC (https://www.server-world.info/query?os=Other&p=vbs&f=1)
do while WScript.Arguments.Count = 0 and WScript.Version >= 5.7
    Dim App
    Dim OS, Value
    Set App = WScript.CreateObject("Shell.Application")
    '##### Check if it is WScript 5.7 or Vista or later
    'Set WMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set OS = Wmi.ExecQuery("SELECT *FROM Win32_OperatingSystem")
    For Each Value in OS
    If left(Value.Version, 3) < 6.0 Then Exit Do
    Next

    '##### Run as administrator.
    App.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ uac", "", "runas"

    WScript.Quit
loop

Dim Fso, Sh

Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
Set Sh = WScript.CreateObject("WScript.Shell")

Dim backupDir, sysRootDir, softDist, bkSoftDist, catRoot, bkCatRoot

backupDir = fso.getParentFolderName(WScript.ScriptFullName) & "\backups"
sysRootDir = Sh.ExpandEnvironmentStrings("%SystemRoot%")
If DebugMode Then sysRootDir = fso.getParentFolderName(WScript.ScriptFullName) & "\Windows"

softDist = sysRootDir & "\SoftwareDistribution"
catRoot = sysRootDir & "\System32\catroot2"
bkSoftDist = backupDir & "\SoftwareDistribution.old"
bkCatRoot = backupDir & "\catroot2.old"

Dim rOrigin, rRestore
Dim nfOrigin(), nfRestore()
ReDim nfOrigin(-1)
ReDim nfRestore(-1)

If DebugMode Then MsgBox szDebugMode, vbInformation Or vbSystemModal, szInfo

rOrigin = CheckOrigin(nfOrigin)
rRestore = CheckRestores(nfRestore)

' Create backup dir
If Not Fso.FolderExists(backupDir) Then
    Fso.CreateFolder(backupDir)
    'Dim fo
    'Set fo = Fso.GetFolder(backupDir)
    'fo.attributes = 2 ' hidden directory
End If

If Not Fso.FolderExists(sysRootDir) Then
    MsgBox szSysRootNotFound + vbCrLf + sysRootDir, vbCritical Or vbSystemModal, szError
    WScript.Quit(1)
End If

If rOrigin <> 2 And rRestore <> 2 Then
    Dim notFounds
    If rOrigin < 2 Then
        notFounds = Join(nfRestore, vbCrLf)
    Else
        notFounds = Join(nfOrigin, vbCrLf)
    End If
    MsgBox szOriginNotFound + vbCrLf + notFounds, vbCritical Or vbSystemModal, szError
    WScript.Quit(1)
End If

Select Case rRestore
    Case 0     ' No backup
        DoBackup()
    Case 2     ' Detect complete backups (2 directories)
        DoRestoreOrDelete()
    Case Else  ' Detect incomplete backups
        BackupFileCheckError(nfRestore)
End Select

' Unhandled...
WScript.Quit(1)


'--- Internal functions ------------------------------------

Sub Append(ByRef arr, ByRef item)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = item
End Sub

Function CheckCommon(ByVal sd, ByVal cr, ByRef notFounds)
    CheckCommon = 0
    If Fso.FolderExists(sd) Then
        CheckCommon = CheckCommon + 1
    Else
        Append notFounds, sd
    End If
    
    If Fso.FolderExists(cr) Then
        CheckCommon = CheckCommon + 1
    Else
        Append notFounds, cr
    End If
End Function

Function CheckOrigin(ByRef notFounds)
    CheckOrigin = CheckCommon(softDist, catRoot, notFounds)
End Function

Function CheckRestores(ByRef notFounds)
    CheckRestores = CheckCommon(bkSoftDist, bkCatRoot, notFounds)
End Function

Sub BackupFileCheckError(ByRef notFounds)
    Dim nfs
    nfs = Join(notFounds, vbCrLf)
    MsgBox szBackupBroken & vbCrLf & nfs, vbCritical Or vbSystemModal, szError
    WScript.Quit(1)
End Sub

Sub DoBackup()
    Dim rDoBackup
    rDoBackup = MsgBox(szAskInit, vbQuestion Or vbSystemModal Or vbYesNo, szQuestion)
    If rDoBackup = vbYes Then
        BackupMain
    Else
        Canceled
    End If
    WScript.Quit(0)
End Sub

Sub DoRestoreOrDelete()
    Dim r
    r = MsgBox(szAskRestore, vbQuestion Or vbSystemModal Or vbYesNo, szQuestion)
    If r = vbYes Then
        RestoreMain
        WScript.Quit(0)
    End If

    r = MsgBox(szAskDelete, vbQuestion Or vbSystemModal Or vbYesNo, szQuestion)
    If r = vbYes Then
        DeleteMain
    Else
        Canceled
    End If
        WScript.Quit(0)
End Sub

Sub Canceled()
    MsgBox szCanceled, vbInformation Or vbSystemModal, szInfo
    WScript.Quit(0)
End Sub

Sub BackupMain()
    Dim r, failedControls()
    ReDim failedControls(-1)
    r = MoveCommon(softDist, catRoot, bkSoftDist, bkCatRoot, False, failedControls)
    If r <> 0 And r <> 3 Then
        Select Case r
            Case 2  ' Service control error
                MsgBox szServiceFailed & vbCrLf & Join(failedControls, vbCrLf), vbCritical Or vbSystemModal, szError
            Case Else
                MsgBox szFailed, vbCritical Or vbSystemModal, szError
        End Select
    Else
        If r = 3 Then MsgBox szServiceStartFailed & vbCrLf & Join(failedControls, vbCrLf), vbExclamation Or vbSystemModal, szError
        Dim sr
        sr = MsgBox(szAskReboot, vbQuestion Or vbSystemModal Or vbYesNo, szQuestion)
        If sr = vbYes  Then
            If Not DebugMode Then Sh.Run "shutdown -r -t 0", 0
        Else
            MsgBox szSucceeded, vbInformation Or vbSystemModal, szInfo
        End If
    End If
    WScript.Quit(r)
End Sub

Sub RestoreMain()
    Dim r, failedControls()
    ReDim failedControls(-1)
    r = MoveCommon(bkSoftDist, bkCatRoot, softDist, catRoot, True, failedControls)
    If r <> 0 And r <> 3 Then
        Select Case r
            Case 2  ' Service control error (Stop & Start)
                MsgBox szServiceFailed & vbCrLf & Join(failedControls, vbCrLf), vbCritical Or vbSystemModal, szError
            Case Else
                MsgBox szFailed, vbCritical Or vbSystemModal, szError
        End Select
    Else
        If r = 3 Then MsgBox szServiceStartFailed & vbCrLf & Join(failedControls, vbCrLf), vbExclamation Or vbSystemModal, szError
        Dim sr
        sr = MsgBox(szAskReboot, vbQuestion Or vbSystemModal Or vbYesNo, szQuestion)
        If sr = vbYes Then
            If Not DebugMode Then Sh.Run "shutdown -r -t 0", 0
        Else
            MsgBox szSucceeded, vbInformation Or vbSystemModal, szInfo
        End If
    End If
    WScript.Quit(r)
End Sub

Sub DeleteMain()
    Dim r
    r = MsgBox(szAskRealy, vbExclamation Or vbSystemModal Or vbYesNo, szQuestion)
    If r = vbNo Then Canceled
    If Fso.FolderExists(bkSoftDist) Then Fso.DeleteFolder bkSoftDist, True
    If Fso.FolderExists(bkCatRoot) Then Fso.DeleteFolder bkCatRoot, True
    MsgBox szSucceeded, vbInformation Or vbSystemModal, szInfo
    WScript.Quit(0)
End Sub

Function MoveCommon(ByRef fromSd, ByRef fromCr, ByRef toSb, ByRef toCr, ByVal Force, ByRef failedControls)
    MoveCommon = 0
    ' exists check
    If Fso.FolderExists(toSb) Or Fso.FolderExists(toCr) Then
        If Not Force Then
            MsgBox szExistsDestDir, vbCritical Or vbSystemModal, szError
            MoveCommon = 1
            Exit Function
        End If
    End If

    Dim r
    ' stop services
    r = StopServices(failedControls)
    If r <> 0 Then
        StartServices(failedControls)
        MoveCommon = 2
        Exit Function
    End If
    WScript.Sleep 200

    ' move
    If Fso.FolderExists(toSb) Then Fso.DeleteFolder toSb, True
    If Fso.FolderExists(toCr) Then Fso.DeleteFolder toCr, True
    Fso.MoveFolder fromSd, toSb
    Fso.MoveFolder fromCr, toCr
    WScript.Sleep 200

    ' restart services
    r = StartServices(failedControls)
    If r <> 0 Then MoveCommon = 3
End Function

Function StopServices(ByRef failedControls)
    StopServices = ServiceControl("StopService", failedControls)
End Function

Function StartServices(ByRef failedControls)
    StartServices = ServiceControl("StartService", failedControls)
End Function

Function ServiceControl(ByRef control, ByRef failedControls)
    Dim DependentServices
    DependentServices = Array("BITS", "CryptSvc", "msiserver", "wuauserv")
    Dim r, svc
    r = 0
    If Not DebugMode Then
        For Each svc In DependentServices
            r = r + ServiceControlCore(control, svc, failedControls)
        Next
    End If
    ServiceControl = r
End Function

Function ServiceControlCore(ByRef control, ByRef service, ByRef failedControls)
    Dim r, retry
    ServiceControlCore = 0
    retry = 0
    Do
        Set r = Wmi.ExecMethod("Win32_Service.Name='" & service & "'", control)
        retry = retry + 1
        WScript.Sleep 200
    Loop While r.ReturnValue <> Accepted And r.ReturnValue = RequestCannotBeSent And retry < RetryMax
    If r.ReturnValue <> Accepted And r.ReturnValue <> ServiceNotBeenStarted And r.ReturnValue <> ServiceAlreadyRunning Then
        Append failedControls, service & " (" & control & ") : ErrorCode=" & r.ReturnValue
        ServiceControlCore = 1
    End If
End Function

' ServiceControl Const
Const Accepted = 0
Const ServiceNotBeenStarted = 6
Const ServiceAlreadyRunning = 10
Const RequestCannotBeSent = 10
Const RetryMax = 5

' Messages
Const szInfo = "情報"
Const szWarning = "警告"
Const szError = "エラー"
Const szQuestion = "選択してください"
Const szDebugMode = "デバッグモードで実行します"
Const szSucceeded = "正常に処理を完了しました"
Const szCanceled = "処理がキャンセルされました"
Const szSysRootNotFound = "システムフォルダが見つかりません。"
Const szOriginNotFound = "必要なフォルダが見つかりません"
Const szBackupBroken = "バックアップファイルの構成が壊れています。 以下のフォルダが見つかりません。"
Const szAskReboot = "更新を反映するにはコンピュータの再起動が必要です。 今すぐ再起動しますか？"
Const szAskInit = "Windows Update に関するファイルをバックアップした上で初期化しますか？"
Const szAskRestore = "Windows Update に関するフォルダのバックアップが見つかりました。 復元を行いますか？ (※現在の構成は破棄されます)"
Const szAskDelete = "Windows Update に関するフォルダのバックアップを削除しますか？ (※削除したファイルは元に戻せません)"
Const szAskRealy = "この処理を実行すると、元に戻すことができなくなります。 本当に処理を実行してもよろしいですか？"
Const szFailed = "不明なエラーにより、処理に失敗しました"
Const szServiceFailed = "依存サービスの停止処理に失敗しました"
Const szServiceStartFailed = "依存サービスの再開処理に失敗しました、手動で再開してください。"
Const szExistsDestDir = "リネーム先のフォルダが既に存在します"
