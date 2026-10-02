Attribute VB_Name = "INST_03_Parser"
Option Explicit

'@ENGINE: VBA Develop Boosters - INST
'@VERSION: 1.0
'@MODULE: INST_03_Parser

'==================================================
' コード解析
'==================================================


'==================================================
' Markdownコードフェンスを除去
'
' ```vb
' コード
' ```
'==================================================
Public Function RemoveCodeFence( _
    ByVal codeText As String) As String

    Dim lines() As String
    Dim resultText As String
    Dim i As Long
    Dim startLine As Long
    Dim endLine As Long

    codeText = Replace(codeText, vbCrLf, vbLf)
    codeText = Replace(codeText, vbCr, vbLf)

    lines = Split(codeText, vbLf)

    startLine = LBound(lines)
    endLine = UBound(lines)

    '----------------------------------------------
    ' 先頭の ```
    '----------------------------------------------
    If Left$(Trim$(lines(startLine)), 3) = "```" Then

        startLine = startLine + 1

    End If

    '----------------------------------------------
    ' 最後の ```
    '----------------------------------------------
    If endLine >= startLine Then

        If Left$(Trim$(lines(endLine)), 3) = "```" Then

            endLine = endLine - 1

        End If

    End If

    '----------------------------------------------
    ' コード再構築
    '----------------------------------------------
    For i = startLine To endLine

        resultText = resultText & lines(i)

        If i < endLine Then

            resultText = resultText & vbCrLf

        End If

    Next i

    RemoveCodeFence = resultText

End Function


'==================================================
' @MODULE ヘッダーからモジュール名を取得
'
' 例：
' '@MODULE: ScriptEngine
'==================================================
Public Function GetModuleName( _
    ByVal codeText As String) As String

    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim moduleText As String

    codeText = Replace(codeText, vbCrLf, vbLf)
    codeText = Replace(codeText, vbCr, vbLf)

    lines = Split(codeText, vbLf)

    For i = LBound(lines) To UBound(lines)

        lineText = Trim$(lines(i))

        '------------------------------------------
        ' コメント形式
        ' '@MODULE:
        '------------------------------------------
        If UCase$(Left$(lineText, 9)) = "'@MODULE:" Then

            moduleText = Mid$(lineText, 10)

            moduleText = Trim$(moduleText)

            GetModuleName = moduleText

            Exit Function

        End If

        '------------------------------------------
        ' コメントなし
        ' @MODULE:
        '------------------------------------------
        If UCase$(Left$(lineText, 8)) = "@MODULE:" Then

            moduleText = Mid$(lineText, 9)

            moduleText = Trim$(moduleText)

            GetModuleName = moduleText

            Exit Function

        End If

    Next i

End Function

'==================================================
' @SUB ヘッダーからプロシージャ名を取得
'
' 例：
' '@SUB: HelloEngine
'==================================================
'==================================================
' @SUB ヘッダーからプロシージャ名を取得
'
' 例：
' '@SUB: HelloEngine
'==================================================
Public Function GetSubName( _
    ByVal codeText As String) As String

    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim subText As String

    codeText = Replace(codeText, vbCrLf, vbLf)
    codeText = Replace(codeText, vbCr, vbLf)

    lines = Split(codeText, vbLf)

    For i = LBound(lines) To UBound(lines)

        lineText = Trim$(lines(i))

        '------------------------------------------
        ' コメント形式
        ' '@SUB:
        '------------------------------------------
        If UCase$(Left$(lineText, 6)) = "'@SUB:" Then

            subText = Mid$(lineText, 7)

            subText = Trim$(subText)

            GetSubName = subText

            Exit Function

        End If

        '------------------------------------------
        ' コメントなし
        ' @SUB:
        '------------------------------------------
        If UCase$(Left$(lineText, 5)) = "@SUB:" Then

            subText = Mid$(lineText, 6)

            subText = Trim$(subText)

            GetSubName = subText

            Exit Function

        End If

    Next i

End Function

