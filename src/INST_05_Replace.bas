Attribute VB_Name = "INST_05_Replace"
Option Explicit

'@ENGINE: VBA Develop Boosters - INST
'@VERSION: 1.0
'@MODULE: INST_05_Replace

'==================================================
' コードリプレイス
'
' @MODULE で対象モジュールを指定
' @SUB    で対象プロシージャを指定
'
' 旧コードは削除せず、
' コメントアウトして履歴として残す。
'==================================================


'==================================================
' クリップボードからプロシージャを置き換える
'==================================================
Public Sub ReplaceProcedureFromClipboard()

    Dim codeText As String
    Dim moduleName As String
    Dim subName As String

    Dim component As Object
    Dim codeModule As Object

    Dim startLine As Long
    Dim lineCount As Long

    Dim oldCode As String
    Dim oldComment As String
    Dim newCode As String

    Dim replacementText As String

    Debug.Print "========================================"
    Debug.Print INST_TITLE
    Debug.Print "Mode : CODE REPLACE"
    Debug.Print "========================================"

    '----------------------------------------------
    ' 1. クリップボード取得
    '----------------------------------------------
    codeText = GetClipboardUnicodeText()

    If Len(codeText) = 0 Then

        MsgBox "クリップボードからコードを取得できませんでした。", _
               vbExclamation

        Exit Sub

    End If

    Debug.Print "Clipboard OK"

    '----------------------------------------------
    ' 2. Markdownコードフェンス除去
    '----------------------------------------------
    codeText = RemoveCodeFence(codeText)

    '----------------------------------------------
    ' 3. @MODULE取得
    '----------------------------------------------
    moduleName = GetModuleName(codeText)

    If Len(moduleName) = 0 Then

        MsgBox "モジュールヘッダーが見つかりません。" & vbCrLf & vbCrLf & _
               "'@MODULE: TestEngine" & vbCrLf & _
               "のようなヘッダーを付けてください。", _
               vbExclamation

        Exit Sub

    End If

    Debug.Print "Module Name : " & moduleName

    '----------------------------------------------
    ' 4. @SUB取得
    '----------------------------------------------
    subName = GetSubName(codeText)

    If Len(subName) = 0 Then

        MsgBox "プロシージャヘッダーが見つかりません。" & vbCrLf & vbCrLf & _
               "'@SUB: HelloEngine" & vbCrLf & _
               "のようなヘッダーを付けてください。", _
               vbExclamation

        Exit Sub

    End If

    Debug.Print "Sub Name    : " & subName

    '----------------------------------------------
    ' 5. 対象モジュール検索
    '----------------------------------------------
    Set component = FindModule(moduleName)

    If component Is Nothing Then

        MsgBox "対象モジュールが見つかりません。" & vbCrLf & vbCrLf & _
               "Module : " & moduleName, _
               vbExclamation

        Debug.Print "Module not found."

        Exit Sub

    End If

    Debug.Print "Module Found"

    '----------------------------------------------
    ' 6. CodeModule取得
    '----------------------------------------------
    Set codeModule = component.codeModule

    '----------------------------------------------
    ' 7. 対象プロシージャ検索
    '----------------------------------------------
    On Error GoTo ProcedureNotFound

    startLine = codeModule.ProcStartLine( _
                    subName, _
                    VBEXT_PK_PROC)

    lineCount = codeModule.ProcCountLines( _
                    subName, _
                    VBEXT_PK_PROC)

    On Error GoTo 0

    Debug.Print "Start Line  : " & startLine
    Debug.Print "Line Count  : " & lineCount

    '----------------------------------------------
    ' 8. 旧コード取得
    '----------------------------------------------
    oldCode = codeModule.lines( _
                  startLine, _
                  lineCount)

    Debug.Print "Old Code    : OK"

    '----------------------------------------------
    ' 9. 新コード取得
    '
    ' ヘッダー部分を除き、
    ' Public Sub ～ End Sub のコードだけを取得
    '----------------------------------------------
    newCode = ExtractProcedureCode(codeText, subName)

    If Len(newCode) = 0 Then

        MsgBox "置き換えるプロシージャ本体を取得できませんでした。" & _
               vbCrLf & vbCrLf & _
               "Sub : " & subName, _
               vbExclamation

        Exit Sub

    End If

    Debug.Print "New Code    : OK"

    '----------------------------------------------
    ' 10. 旧コードをコメント化
    '----------------------------------------------
    oldComment = CommentOutCode(oldCode)

    '----------------------------------------------
    ' 11. 置換文字列作成
    '----------------------------------------------
    replacementText = _
        "'--- [INST REPLACED] " & _
        Format$(Date, "yyyy/mm/dd") & _
        " --------------------------------" & vbCrLf & _
        oldComment & _
        vbCrLf & _
        "'--- [END OLD CODE] --------------------------------" & _
        vbCrLf & _
        vbCrLf & _
        newCode

    '----------------------------------------------
    ' 12. 旧プロシージャを置換
    '
    ' 旧コードそのものはコメント化して
    ' 同じ場所へ保存する。
    '----------------------------------------------
    codeModule.DeleteLines _
        startLine, _
        lineCount

    codeModule.InsertLines _
        startLine, _
        replacementText

    Debug.Print "Replaced    : " & subName
    Debug.Print "Old Code    : preserved"
    Debug.Print "========================================"

    MsgBox "プロシージャを置き換えました。" & vbCrLf & vbCrLf & _
           "Module : " & moduleName & vbCrLf & _
           "Sub    : " & subName & vbCrLf & vbCrLf & _
           "旧コードはコメントとして保存されています。", _
           vbInformation

    Exit Sub


'==================================================
' プロシージャが見つからない
'==================================================
ProcedureNotFound:

    On Error GoTo 0

    Debug.Print "Procedure not found : " & subName

    MsgBox "対象プロシージャが見つかりません。" & vbCrLf & vbCrLf & _
           "Module : " & moduleName & vbCrLf & _
           "Sub    : " & subName, _
           vbExclamation

End Sub


'==================================================
' モジュール検索
'==================================================
Private Function FindModule( _
    ByVal moduleName As String) As Object

    Dim component As Object

    Set FindModule = Nothing

    For Each component In ThisWorkbook.VBProject.VBComponents

        If StrComp(component.Name, _
                   moduleName, _
                   vbTextCompare) = 0 Then

            Set FindModule = component

            Exit Function

        End If

    Next component

End Function


'==================================================
' プロシージャコードを取得
'
' ヘッダー
' Option Explicit
' その他のコード
'
' を除き、
'
' Public Sub ～ End Sub
'
' の部分だけを取得する。
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

    codeText = Replace(codeText, vbCrLf, vbLf)
    codeText = Replace(codeText, vbCr, vbLf)

    lines = Split(codeText, vbLf)

    startIndex = -1
    endIndex = -1

    '----------------------------------------------
    ' プロシージャ開始位置検索
    '----------------------------------------------
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

    '----------------------------------------------
    ' End Sub / End Function / End Property
    '----------------------------------------------
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

    '----------------------------------------------
    ' コード再構築
    '----------------------------------------------
    For i = startIndex To endIndex

        resultText = resultText & lines(i)

        If i < endIndex Then

            resultText = resultText & vbCrLf

        End If

    Next i

    ExtractProcedureCode = resultText

End Function


'==================================================
' 対象プロシージャ開始行判定
'==================================================
Private Function IsTargetProcedureStart( _
    ByVal lineText As String, _
    ByVal subName As String) As Boolean

    Dim upperName As String

    upperName = UCase$(subName)

    If lineText = "PUBLIC SUB " & upperName & "()" _
    Or lineText = "PRIVATE SUB " & upperName & "()" _
    Or lineText = "SUB " & upperName & "()" Then

        IsTargetProcedureStart = True

        Exit Function

    End If

    If lineText = "PUBLIC FUNCTION " & upperName & "()" _
    Or lineText = "PRIVATE FUNCTION " & upperName & "()" _
    Or lineText = "FUNCTION " & upperName & "()" Then

        IsTargetProcedureStart = True

        Exit Function

    End If

End Function


'==================================================
' コードをコメント化
'
' 元のコードを一行ずつ
'
' '元のコード
'
' の形式にする。
'==================================================
Private Function CommentOutCode( _
    ByVal codeText As String) As String

    Dim lines() As String
    Dim i As Long

    Dim resultText As String

    codeText = Replace(codeText, vbCrLf, vbLf)
    codeText = Replace(codeText, vbCr, vbLf)

    lines = Split(codeText, vbLf)

    For i = LBound(lines) To UBound(lines)

        resultText = resultText & _
                     "'" & lines(i)

        If i < UBound(lines) Then

            resultText = resultText & vbCrLf

        End If

    Next i

    CommentOutCode = resultText

End Function
