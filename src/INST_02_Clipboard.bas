Attribute VB_Name = "INST_02_Clipboard"
Option Explicit

'@ENGINE: VBA Develop Boosters - INST
'@VERSION: 1.0
'@MODULE: INST_02_Clipboard

'==================================================
' Windows API
' Excel 2007 / VBA 6.5 / 32bit 用
' PtrSafe は使用しない
'==================================================

Private Declare Function OpenClipboard Lib "user32" ( _
    ByVal hWnd As Long) As Long

Private Declare Function CloseClipboard Lib "user32" () As Long

Private Declare Function GetClipboardData Lib "user32" ( _
    ByVal uFormat As Long) As Long

Private Declare Function GlobalLock Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Private Declare Function GlobalUnlock Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" ( _
    ByVal lpString As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)


'==================================================
' クリップボードからUnicode文字列を取得
'==================================================
Public Function GetClipboardUnicodeText() As String

    Dim hData As Long
    Dim lpData As Long
    Dim textLength As Long
    Dim resultText As String

    If OpenClipboard(0&) = 0 Then

        Debug.Print "OpenClipboard NG"

        Exit Function

    End If

    hData = GetClipboardData(INST_CF_UNICODETEXT)

    If hData = 0 Then

        Debug.Print "CF_UNICODETEXT がありません"

        CloseClipboard

        Exit Function

    End If

    lpData = GlobalLock(hData)

    If lpData = 0 Then

        Debug.Print "GlobalLock NG"

        CloseClipboard

        Exit Function

    End If

    textLength = lstrlenW(lpData)

    If textLength > 0 Then

        resultText = String$(textLength, vbNullChar)

        CopyMemory ByVal StrPtr(resultText), _
                   ByVal lpData, _
                   textLength * 2

    End If

    GlobalUnlock hData

    CloseClipboard

    GetClipboardUnicodeText = resultText

End Function
