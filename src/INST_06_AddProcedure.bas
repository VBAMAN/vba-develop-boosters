Attribute VB_Name = "INST_06_AddProcedure"
Option Explicit

'@ENGINE: VBA Develop Boosters - INST
'@VERSION: 1.0
'@MODULE: INST_06_AddProcedure

'==================================================
' プロシージャ追加
'
' @MODULE で対象モジュールを指定
' @SUB    で追加するプロシージャを指定
'
' 既存プロシージャが存在する場合は追加しない。
'
' INST_05_Replace とは明確に役割を分離する。
'
'   INST_05_Replace
'       → 既存プロシージャを置き換える
'
'   INST_06_AddProcedure
'       → 新しいプロシージャを追加する
'
'==================================================

Public Sub AddProcedureFromClipboard()

    Dim codeText As String
    Dim moduleName As String
    Dim subName As String

    Dim component As Object
    Dim codeModule As Object

    Dim newCode As String
    Dim insertLine As Long

    Debug.Print "========================================"
    Debug.Print INST_TITLE
    Debug.Print "Mode : ADD PROCEDURE"
    Debug.Print "========================================"

    '------------------------------------------
    ' クリップボード取得
    '------------------------------------------

    codeText = GetClipboardUnicodeText()

    If Len(codeText) = 0 Then

        MsgBox "クリップボードからコードを取得できませんでした。", _
               vbExclamation

        Exit Sub

    End If

    Debug.Print "Clipboard OK"

    '------------------------------------------
    ' コードフェンス除去
    '------------------------------------------

    codeText = RemoveCodeFence(codeText)

    '------------------------------------------
    ' モジュール名取得
    '------------------------------------------

    moduleName = GetModuleName(codeText)

    If Len(moduleName) = 0 Then

        MsgBox "モジュールヘッダーが見つかりません。" & _
               vbCrLf & vbCrLf & _
               "'@MODULE: TestEngine" & _
               vbCrLf & _
               "のようなヘッダーを付けてください。", _
               vbExclamation

        Exit Sub

    End If

    Debug.Print "Module Name : " & moduleName

    '------------------------------------------
    ' プロシージャ名取得
    '------------------------------------------

    subName = GetSubName(codeText)

    If Len(subName) = 0 Then

        MsgBox "プロシージャヘッダーが見つかりません。" & _
               vbCrLf & vbCrLf & _
               "'@SUB: HelloEngine" & _
               vbCrLf & _
               "のようなヘッダーを付けてください。", _
               vbExclamation

        Exit Sub

    End If

    Debug.Print "Sub Name    : " & subName

    '------------------------------------------
    ' 対象モジュール検索
    '------------------------------------------

    Set component = FindModule(moduleName)

    If component Is Nothing Then

        MsgBox "対象モジュールが見つかりません。" & _
               vbCrLf & vbCrLf & _
               "Module : " & moduleName, _
               vbExclamation

        Debug.Print "Module not found."

        Exit Sub

    End If

    Debug.Print "Module Found"

    Set codeModule = component.codeModule

    '------------------------------------------
    ' 既存プロシージャ確認
    '
    ' ADDモードなので、
    ' すでに存在していたら処理しない。
    '------------------------------------------

    If ProcedureExists( _
            codeModule, _
            subName) Then

        MsgBox "同名のプロシージャがすでに存在します。" & _
               vbCrLf & vbCrLf & _
               "Module : " & moduleName & _
               vbCrLf & _
               "Sub    : " & subName & _
               vbCrLf & vbCrLf & _
               "既存コードは変更していません。", _
               vbExclamation

        Debug.Print "Procedure already exists : " & subName

        Exit Sub

    End If

    Debug.Print "Procedure does not exist."

    '------------------------------------------
    ' 追加するプロシージャ取得
    '------------------------------------------

    newCode = ExtractProcedureCode( _
                  codeText, _
                  subName)

    If Len(newCode) = 0 Then

        MsgBox "追加するプロシージャ本体を取得できませんでした。" & _
               vbCrLf & vbCrLf & _
               "Sub : " & subName, _
               vbExclamation

        Exit Sub

    End If

    Debug.Print "New Code    : OK"

    '------------------------------------------
    ' モジュール末尾へ追加
    '------------------------------------------

    insertLine = codeModule.CountOfLines + 1

    If insertLine > 1 Then

        codeModule.InsertLines _
            insertLine, _
            vbCrLf & newCode

    Else

        codeModule.InsertLines _
            insertLine, _
            newCode

    End If

    Debug.Print "Added       : " & subName
    Debug.Print "========================================"

    MsgBox "プロシージャを追加しました。" & _
           vbCrLf & vbCrLf & _
           "Module : " & moduleName & _
           vbCrLf & _
           "Sub    : " & subName, _
           vbInformation

End Sub


'==================================================
' プロシージャ存在確認
'==================================================

Private Function ProcedureExists( _
    ByVal codeModule As Object, _
    ByVal subName As String) As Boolean

    Dim startLine As Long

    On Error GoTo NotFound

    startLine = codeModule.ProcStartLine( _
                    subName, _
                    VBEXT_PK_PROC)

    ProcedureExists = True

    Exit Function

NotFound:

    ProcedureExists = False

End Function


'==================================================
' プロシージャ本体抽出
'
' @SUB で指定されたプロシージャを
' クリップボードコードから取り出す。
'==================================================

Private Function ExtractProcedureCode( _
    ByVal codeText As String, _
    ByVal subName As String) As String

    Dim lines() As String

    Dim i As Long
    Dim lineText As String
    Dim normalizedText As String

    Dim startIndex As Long
    Dim endIndex As Long

    Dim resultText As String

    codeText = Replace( _
                   codeText, _
                   vbCrLf, _
                   vbLf)

    codeText = Replace( _
                   codeText, _
                   vbCr, _
                   vbLf)

    lines = Split(codeText, vbLf)

    startIndex = -1
    endIndex = -1

    '------------------------------------------
    ' 開始行検索
    '------------------------------------------

    For i = LBound(lines) To UBound(lines)

        lineText = Trim$(lines(i))

        normalizedText = UCase$(lineText)

        If IsTargetProcedureStart( _
                normalizedText, _
                subName) Then

            startIndex = i

            Exit For

        End If

    Next i

    If startIndex < 0 Then Exit Function

    '------------------------------------------
    ' 終了行検索
    '------------------------------------------

    For i = startIndex To UBound(lines)

        lineText = UCase$(Trim$(lines(i)))

        If lineText = "END SUB" _
        Or lineText = "END FUNCTION" _
        Or lineText = "END PROPERTY" Then

            endIndex = i

            Exit For

        End If

    Next i

    If endIndex < 0 Then Exit Function

    '------------------------------------------
    ' プロシージャ取得
    '------------------------------------------

    For i = startIndex To endIndex

        resultText = resultText & lines(i)

        If i < endIndex Then

            resultText = resultText & vbCrLf

        End If

    Next i

    ExtractProcedureCode = resultText

End Function


'==================================================
' プロシージャ開始行判定
'==================================================

Private Function IsTargetProcedureStart( _
    ByVal lineText As String, _
    ByVal subName As String) As Boolean

    Dim upperName As String

    upperName = UCase$(subName)

    '------------------------------------------
    ' Sub
    '------------------------------------------

    If lineText = "PUBLIC SUB " & upperName & "()" _
    Or lineText = "PRIVATE SUB " & upperName & "()" _
    Or lineText = "SUB " & upperName & "()" Then

        IsTargetProcedureStart = True

        Exit Function

    End If

    '------------------------------------------
    ' Function
    '------------------------------------------

    If lineText = "PUBLIC FUNCTION " & upperName & "()" _
    Or lineText = "PRIVATE FUNCTION " & upperName & "()" _
    Or lineText = "FUNCTION " & upperName & "()" Then

        IsTargetProcedureStart = True

        Exit Function

    End If

End Function

'==================================================
' モジュール検索
'==================================================

Private Function FindModule( _
    ByVal moduleName As String) As Object

    Dim component As Object

    Set FindModule = Nothing

    For Each component In ThisWorkbook.VBProject.VBComponents

        If StrComp( _
                component.Name, _
                moduleName, _
                vbTextCompare) = 0 Then

            Set FindModule = component

            Exit Function

        End If

    Next component

End Function

Public Sub SelfTestAdded()

    Debug.Print "INST_06_AddProcedure に自分自身で追加されました。"

End Sub
