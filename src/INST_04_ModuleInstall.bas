Attribute VB_Name = "INST_04_ModuleInstall"
Option Explicit

'@ENGINE: VBA Develop Boosters - INST
'@VERSION: 1.0
'@MODULE: INST_04_ModuleInstall

'==================================================
' モジュールインストール
'
' このモジュールでは、標準モジュールを
' このブックへ組み込む。
'
' ■ クリップボードから組み込む
'     InstallModuleFromClipboard
'
' ■ 完成済み.basファイルを組み込む
'     InstallModuleFromBas
'
' ■ 他のExcelブックから組み込む
'     InstallModuleFromWorkbook
'
' 3つの入口は、最終的に
' ImportBasFile
' を使用してExcelへImportする。
'
'==================================================


'==================================================
' クリップボードからExcelへモジュール追加
'==================================================

Public Sub InstallModuleFromClipboard()

    Dim codeText As String
    Dim moduleName As String
    Dim tempPath As String

    Debug.Print "========================================"
    Debug.Print INST_TITLE
    Debug.Print "Mode : MODULE INSTALL / CLIPBOARD"
    Debug.Print "========================================"

    '----------------------------------------------
    ' 1. クリップボード取得
    '----------------------------------------------

    codeText = GetClipboardUnicodeText()

    If Len(codeText) = 0 Then

        MsgBox _
            "クリップボードからコードを取得できませんでした。", _
            vbExclamation

        Exit Sub

    End If

    Debug.Print "Clipboard OK"

    '----------------------------------------------
    ' 2. Markdownコードフェンス除去
    '----------------------------------------------

    codeText = RemoveCodeFence(codeText)

    '----------------------------------------------
    ' 3. @MODULEからモジュール名取得
    '----------------------------------------------

    moduleName = GetModuleName(codeText)

    If Len(moduleName) = 0 Then

        MsgBox _
            "モジュールヘッダーが見つかりません。" & vbCrLf & vbCrLf & _
            "コード先頭に" & vbCrLf & _
            "'@MODULE: ScriptEngine" & vbCrLf & _
            "のようなヘッダーを付けてください。", _
            vbExclamation

        Exit Sub

    End If

    Debug.Print "Module Name : " & moduleName

    '----------------------------------------------
    ' 4. 同名モジュール確認
    '----------------------------------------------

    If ModuleExists(moduleName) Then

        MsgBox _
            "同名のモジュールがすでに存在します。" & vbCrLf & vbCrLf & _
            "Module : " & moduleName, _
            vbExclamation

        Debug.Print "Module already exists."

        Exit Sub

    End If

    '----------------------------------------------
    ' 5. 一時.basファイル作成
    '----------------------------------------------

    tempPath = CreateTempBasFile( _
                   codeText, _
                   moduleName)

    If Len(tempPath) = 0 Then

        MsgBox _
            ".basファイルの作成に失敗しました。", _
            vbCritical

        Exit Sub

    End If

    Debug.Print "Temp File : " & tempPath

    '----------------------------------------------
    ' 6. 共通Import処理
    '----------------------------------------------

    If ImportBasFile( _
            tempPath, _
            moduleName) = False Then

        DeleteTempFile tempPath

        Exit Sub

    End If

    '----------------------------------------------
    ' 7. 一時ファイル削除
    '----------------------------------------------

    DeleteTempFile tempPath

    Debug.Print "Installed : " & moduleName
    Debug.Print "========================================"

    MsgBox _
        "モジュールを追加しました。" & vbCrLf & vbCrLf & _
        "Module : " & moduleName, _
        vbInformation

End Sub


'==================================================
' 完成済み.basファイルをExcelへ追加
'
' sourceBasPath : .basファイルのフルパス
' newModuleName : 追加後のモジュール名
'
' newModuleNameを空欄にした場合は、
' .basファイル内のモジュール名を使用する。
'==================================================

Public Sub InstallModuleFromBas( _
    ByVal sourceBasPath As String, _
    Optional ByVal newModuleName As String = "")

    Dim moduleName As String

    Debug.Print "========================================"
    Debug.Print INST_TITLE
    Debug.Print "Mode : MODULE INSTALL / BAS"
    Debug.Print "========================================"

    '----------------------------------------------
    ' .basファイル存在確認
    '----------------------------------------------

    If Len(Dir$(sourceBasPath)) = 0 Then

        MsgBox _
            ".basファイルが見つかりません。" & vbCrLf & vbCrLf & _
            sourceBasPath, _
            vbExclamation

        Exit Sub

    End If

    '----------------------------------------------
    ' モジュール名
    '
    ' 指定がなければファイル名から取得
    '----------------------------------------------

    moduleName = newModuleName

    If Len(moduleName) = 0 Then

        moduleName = GetBasModuleName(sourceBasPath)

    End If

    If Len(moduleName) = 0 Then

        MsgBox _
            "モジュール名を取得できませんでした。" & vbCrLf & vbCrLf & _
            "ファイル：" & sourceBasPath, _
            vbExclamation

        Exit Sub

    End If

    Debug.Print "Source BAS  : " & sourceBasPath
    Debug.Print "Module Name : " & moduleName

    '----------------------------------------------
    ' 同名モジュール確認
    '----------------------------------------------

    If ModuleExists(moduleName) Then

        MsgBox _
            "同名のモジュールがすでに存在します。" & vbCrLf & vbCrLf & _
            "Module : " & moduleName, _
            vbExclamation

        Exit Sub

    End If

    '----------------------------------------------
    ' Import
    '----------------------------------------------

    If ImportBasFile( _
            sourceBasPath, _
            moduleName) = False Then

        Exit Sub

    End If

    Debug.Print "Installed : " & moduleName
    Debug.Print "========================================"

    MsgBox _
        "basファイルを取り込みました。" & vbCrLf & vbCrLf & _
        "Module : " & moduleName, _
        vbInformation

End Sub


'==================================================
' 他のExcelブックから標準モジュールを追加
'
' sourceBookPath : 元ブックのフルパス
' sourceModule   : 元ブックのモジュール名
' newModuleName  : このブックでのモジュール名
'
' 元ブックの標準モジュールを一時.basへExportし、
' このブックへImportする。
'==================================================

Public Sub InstallModuleFromWorkbook( _
    ByVal sourceBookPath As String, _
    ByVal sourceModule As String, _
    ByVal newModuleName As String)

    Dim sourceBook As Workbook
    Dim sourceComponent As Object

    Dim tempPath As String

    Dim sourceBookWasOpen As Boolean

    Debug.Print "========================================"
    Debug.Print INST_TITLE
    Debug.Print "Mode : MODULE INSTALL / WORKBOOK"
    Debug.Print "========================================"

    Debug.Print "Source Book   : " & sourceBookPath
    Debug.Print "Source Module : " & sourceModule
    Debug.Print "New Module    : " & newModuleName

    '----------------------------------------------
    ' 元ブックの存在確認
    '----------------------------------------------

    If Len(Dir$(sourceBookPath)) = 0 Then

        MsgBox _
            "元ブックが見つかりません。" & vbCrLf & vbCrLf & _
            sourceBookPath, _
            vbExclamation

        Exit Sub

    End If

    '----------------------------------------------
    ' 追加先モジュール確認
    '----------------------------------------------

    If ModuleExists(newModuleName) Then

        MsgBox _
            "同名のモジュールがすでに存在します。" & _
            vbCrLf & vbCrLf & _
            "Module : " & newModuleName, _
            vbExclamation

        Exit Sub

    End If

    '----------------------------------------------
    ' すでに開いているブックを取得
    '----------------------------------------------

    On Error Resume Next

    Set sourceBook = _
        Workbooks(Dir$(sourceBookPath))

    On Error GoTo ErrorHandler

    If sourceBook Is Nothing Then

        '------------------------------------------
        ' 開いていなければReadOnlyで開く
        '------------------------------------------

        Set sourceBook = _
            Workbooks.Open( _
                fileName:=sourceBookPath, _
                ReadOnly:=True)

        sourceBookWasOpen = False

    Else

        sourceBookWasOpen = True

    End If

    Debug.Print "Source Book Opened"

    '----------------------------------------------
    ' 元モジュール取得
    '----------------------------------------------

    Set sourceComponent = _
        sourceBook.VBProject.VBComponents( _
            sourceModule)

    Debug.Print "Source Module Found"

    '----------------------------------------------
    ' 一時.basファイル作成
    '----------------------------------------------

    tempPath = _
        Environ$("TEMP") & _
        Application.PathSeparator & _
        INST_TEMP_FILE_PREFIX & _
        newModuleName & ".bas"

    '----------------------------------------------
    ' 同名一時ファイル削除
    '----------------------------------------------

    DeleteTempFile tempPath

    '----------------------------------------------
    ' 元モジュールExport
    '----------------------------------------------

    sourceComponent.Export tempPath

    Debug.Print "Exported : " & tempPath

    '----------------------------------------------
    ' 共通Import処理
    '----------------------------------------------

    If ImportBasFile( _
            tempPath, _
            newModuleName) = False Then

        GoTo Cleanup

    End If

    Debug.Print "Installed : " & newModuleName

Cleanup:

    '----------------------------------------------
    ' 一時ファイル削除
    '----------------------------------------------

    DeleteTempFile tempPath

    '----------------------------------------------
    ' この処理で開いたブックだけ閉じる
    '----------------------------------------------

    If Not sourceBook Is Nothing Then

        If sourceBookWasOpen = False Then

            sourceBook.Close _
                SaveChanges:=False

        End If

    End If

    Set sourceComponent = Nothing
    Set sourceBook = Nothing

    Debug.Print "========================================"

    If Len(Dir$(tempPath)) = 0 Then

        If ModuleExists(newModuleName) Then

            MsgBox _
                "モジュールを取り込みました。" & _
                vbCrLf & vbCrLf & _
                "Book   : " & sourceBookPath & _
                vbCrLf & _
                "Module : " & sourceModule & _
                vbCrLf & _
                "New    : " & newModuleName, _
                vbInformation

        End If

    End If

    Exit Sub


ErrorHandler:

    Debug.Print "InstallModuleFromWorkbook Error"
    Debug.Print "Error Number : " & Err.Number
    Debug.Print "Description  : " & Err.Description

    MsgBox _
        "モジュールの取り込みに失敗しました。" & _
        vbCrLf & vbCrLf & _
        "ブック：" & sourceBookPath & _
        vbCrLf & _
        "モジュール：" & sourceModule & _
        vbCrLf & _
        "エラー番号：" & Err.Number & _
        vbCrLf & _
        "内容：" & Err.Description, _
        vbExclamation

    Resume Cleanup

End Sub


'==================================================
' 共通.bas Import
'
' sourceBasPath : Importする.basファイル
' newModuleName : Import後のモジュール名
'
' 戻り値
'   True  = 成功
'   False = 失敗
'==================================================

Private Function ImportBasFile( _
    ByVal sourceBasPath As String, _
    ByVal newModuleName As String) As Boolean

    Dim importedComponent As Object

    ImportBasFile = False

    Debug.Print "Import BAS"
    Debug.Print "File   : " & sourceBasPath
    Debug.Print "Module : " & newModuleName

    '----------------------------------------------
    ' ファイル存在確認
    '----------------------------------------------

    If Len(Dir$(sourceBasPath)) = 0 Then

        MsgBox _
            ".basファイルが見つかりません。" & _
            vbCrLf & vbCrLf & _
            sourceBasPath, _
            vbExclamation

        Exit Function

    End If

    '----------------------------------------------
    ' Import
    '----------------------------------------------

    On Error GoTo ErrorHandler

    Set importedComponent = _
        ThisWorkbook.VBProject.VBComponents.Import( _
            sourceBasPath)

    '----------------------------------------------
    ' Import後のモジュール名変更
    '----------------------------------------------

    importedComponent.Name = newModuleName

    Set importedComponent = Nothing

    ImportBasFile = True

    Debug.Print "Import OK"

    Exit Function


ErrorHandler:

    Debug.Print "ImportBasFile Error : " & Err.Number
    Debug.Print Err.Description

    MsgBox _
        "モジュールの追加に失敗しました。" & _
        vbCrLf & vbCrLf & _
        "File   : " & sourceBasPath & _
        vbCrLf & _
        "Module : " & newModuleName & _
        vbCrLf & _
        "Error  : " & Err.Number & _
        vbCrLf & _
        "内容：" & Err.Description, _
        vbCritical

    Set importedComponent = Nothing

End Function


'==================================================
' 同名モジュール存在確認
'==================================================

Private Function ModuleExists( _
    ByVal moduleName As String) As Boolean

    Dim component As Object

    ModuleExists = False

    For Each component In _
        ThisWorkbook.VBProject.VBComponents

        If StrComp( _
                component.Name, _
                moduleName, _
                vbTextCompare) = 0 Then

            ModuleExists = True

            Exit Function

        End If

    Next component

End Function


'==================================================
' クリップボードコードから
' 一時.basファイルを作成
'
' VBComponents.Import用
'
' 日本語Windows環境ではShift-JISで保存
'==================================================

Private Function CreateTempBasFile( _
    ByVal codeText As String, _
    ByVal moduleName As String) As String

    Dim stream As Object
    Dim tempPath As String

    CreateTempBasFile = ""

    On Error GoTo ErrorHandler

    tempPath = _
        Environ$("TEMP") & _
        Application.PathSeparator & _
        INST_TEMP_FILE_PREFIX & _
        moduleName & ".bas"

    '----------------------------------------------
    ' 既存ファイル削除
    '----------------------------------------------

    DeleteTempFile tempPath

    '----------------------------------------------
    ' ADODB.Stream
    '----------------------------------------------

    Set stream = _
        CreateObject("ADODB.Stream")

    stream.Type = 2

    '----------------------------------------------
    ' VBAの.bas Import用
    ' 日本語WindowsのANSIコードページ
    '----------------------------------------------

    stream.Charset = "shift_jis"

    stream.Open

    stream.WriteText codeText

    '----------------------------------------------
    ' 2 = 上書き保存
    '----------------------------------------------

    stream.SaveToFile _
        tempPath, _
        2

    stream.Close

    Set stream = Nothing

    CreateTempBasFile = tempPath

    Debug.Print _
        "Create BAS OK : " & tempPath

    Exit Function


ErrorHandler:

    Debug.Print _
        "CreateTempBasFile Error : " & Err.Number

    Debug.Print Err.Description

    On Error Resume Next

    If Not stream Is Nothing Then

        stream.Close

    End If

    Set stream = Nothing

    DeleteTempFile tempPath

    CreateTempBasFile = ""

End Function


'==================================================
' .basファイル名からモジュール名を取得
'
' 例：
'   C:\Temp\TestEngine.bas
'
'   ↓
'
'   TestEngine
'
'==================================================

Private Function GetBasModuleName( _
    ByVal basPath As String) As String

    Dim fileName As String
    Dim dotPosition As Long

    GetBasModuleName = ""

    fileName = Dir$(basPath)

    If Len(fileName) = 0 Then Exit Function

    dotPosition = InStrRev( _
                      fileName, _
                      ".")

    If dotPosition > 1 Then

        GetBasModuleName = _
            Left$( _
                fileName, _
                dotPosition - 1)

    Else

        GetBasModuleName = fileName

    End If

End Function


'==================================================
' 一時ファイル削除
'==================================================

Private Sub DeleteTempFile( _
    ByVal filePath As String)

    If Len(filePath) = 0 Then Exit Sub

    On Error Resume Next

    If Dir$(filePath) <> "" Then

        Kill filePath

    End If

    On Error GoTo 0

End Sub

