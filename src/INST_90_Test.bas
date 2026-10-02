Attribute VB_Name = "INST_90_Test"
Option Explicit

'@ENGINE: VBA Develop Boosters - INST
'@VERSION: 1.0
'@MODULE: INST_90_Test

'==================================================
' INST テスト
'
' INST_04_ModuleInstall のテスト用モジュール
'
' ■ Test_ModuleInstall_Clipboard
'     クリップボードからモジュール追加
'
' ■ Test_ModuleInstall_Bas
'     完成済み.basファイルからモジュール追加
'
' ■ Test_ModuleInstall_Workbook
'     他ブックから指定モジュールを追加
'
' ■ Test_ModuleInstall_All
'     3モードの案内
'
'==================================================


'==================================================
' TEST 01
'
' クリップボードからモジュールを追加
'
' 実行前に、以下のようなコードをコピーする。
'
' '@ENGINE: VBA Develop Boosters - INST TEST
' '@VERSION: 1.0
' '@MODULE: INST_TestClipboard
'
' Option Explicit
'
' Public Sub TestClipboardModule()
'
'     Debug.Print "Clipboard Module OK"
'
' End Sub
'
'==================================================

Public Sub Test_ModuleInstall_Clipboard()

    Debug.Print "========================================"
    Debug.Print "TEST : MODULE INSTALL / CLIPBOARD"
    Debug.Print "========================================"

    MsgBox _
        "クリップボードにテスト用モジュールをコピーしてから" & _
        vbCrLf & _
        "［OK］を押してください。" & _
        vbCrLf & vbCrLf & _
        "実行する処理：" & vbCrLf & _
        "InstallModuleFromClipboard", _
        vbInformation

    InstallModuleFromClipboard

End Sub


'==================================================
' TEST 02
'
' 完成済み.basファイルからモジュールを追加
'
' このブックと同じフォルダに
'
' INST_TestBas.bas
'
' を置いて実行する。
'
'==================================================

Public Sub Test_ModuleInstall_Bas()

    Dim folderPath As String
    Dim sourcePath As String

    Debug.Print "========================================"
    Debug.Print "TEST : MODULE INSTALL / BAS"
    Debug.Print "========================================"

    '----------------------------------------------
    ' このブックのフォルダ
    '----------------------------------------------

    folderPath = _
        ThisWorkbook.Path & _
        Application.PathSeparator

    sourcePath = _
        folderPath & _
        "INST_TestBas.bas"

    Debug.Print "BAS File : " & sourcePath

    '----------------------------------------------
    ' ファイル確認
    '----------------------------------------------

    If Dir$(sourcePath) = "" Then

        MsgBox _
            "テスト用.basファイルが見つかりません。" & _
            vbCrLf & vbCrLf & _
            sourcePath & _
            vbCrLf & vbCrLf & _
            "このファイルを配置してから実行してください。", _
            vbExclamation

        Exit Sub

    End If

    '----------------------------------------------
    ' Import
    '----------------------------------------------

    InstallModuleFromBas _
        sourcePath, _
        "INST_TestBas"

End Sub


'==================================================
' TEST 03
'
' 他のExcelブックから標準モジュールを追加
'
' このブックと同じフォルダに
'
' INST_TestSource.xlsm
'
' を置く。
'
' そのブックに
'
' INST_TestSource
'
' という標準モジュールを用意する。
'
'==================================================

Public Sub Test_ModuleInstall_Workbook()

    Dim folderPath As String
    Dim sourcePath As String

    Debug.Print "========================================"
    Debug.Print "TEST : MODULE INSTALL / WORKBOOK"
    Debug.Print "========================================"

    '----------------------------------------------
    ' このブックのフォルダ
    '----------------------------------------------

    folderPath = _
        ThisWorkbook.Path & _
        Application.PathSeparator

    sourcePath = _
        folderPath & _
        "INST_TestSource.xlsm"

    Debug.Print "Source Book : " & sourcePath

    '----------------------------------------------
    ' 元ブック確認
    '----------------------------------------------

    If Dir$(sourcePath) = "" Then

        MsgBox _
            "テスト用Excelブックが見つかりません。" & _
            vbCrLf & vbCrLf & _
            sourcePath & _
            vbCrLf & vbCrLf & _
            "このファイルを配置してから実行してください。", _
            vbExclamation

        Exit Sub

    End If

    '----------------------------------------------
    ' 指定モジュールをImport
    '----------------------------------------------

    InstallModuleFromWorkbook _
        sourcePath, _
        "INST_TestSource", _
        "INST_TestSource_Imported"

End Sub


'==================================================
' TEST ALL
'
' 3つのテスト方法を確認
'==================================================

Public Sub Test_ModuleInstall_All()

    Debug.Print "========================================"
    Debug.Print "INST MODULE INSTALL TEST"
    Debug.Print "========================================"

    Debug.Print ""
    Debug.Print "[01] Clipboard"
    Debug.Print "     Test_ModuleInstall_Clipboard"

    Debug.Print ""
    Debug.Print "[02] BAS File"
    Debug.Print "     Test_ModuleInstall_Bas"

    Debug.Print ""
    Debug.Print "[03] Workbook"
    Debug.Print "     Test_ModuleInstall_Workbook"

    Debug.Print ""
    Debug.Print "========================================"

    MsgBox _
        "INST_04_ModuleInstall のテストです。" & _
        vbCrLf & vbCrLf & _
        "01 : Clipboard" & _
        vbCrLf & _
        "02 : .bas File" & _
        vbCrLf & _
        "03 : Workbook Module" & _
        vbCrLf & vbCrLf & _
        "個別のテストSubを実行してください。", _
        vbInformation

End Sub
