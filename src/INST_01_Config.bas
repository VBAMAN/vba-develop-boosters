Attribute VB_Name = "INST_01_Config"
Option Explicit

'@ENGINE: VBA Develop Boosters - INST
'@VERSION: 1.0
'@MODULE: INST_01_Config

'==================================================
' INST 共通設定
'==================================================

'----------------------------------------------
' クリップボード
'----------------------------------------------
Public Const INST_CF_UNICODETEXT As Long = 13

'----------------------------------------------
' 一時ファイル
'----------------------------------------------
Public Const INST_TEMP_FILE_PREFIX As String = _
    "ExcelPlate_"

'----------------------------------------------
' ヘッダー
'----------------------------------------------
Public Const INST_HEADER_ENGINE As String = _
    "@ENGINE:"

Public Const INST_HEADER_VERSION As String = _
    "@VERSION:"

Public Const INST_HEADER_MODULE As String = _
    "@MODULE:"

Public Const INST_HEADER_SUB As String = _
    "@SUB:"

'----------------------------------------------
' エラー表示
'----------------------------------------------
Public Const INST_TITLE As String = _
    "Excel Plate Module Installer"

'----------------------------------------------
' VBComponent / CodeModule
' vbext_ProcKind
'
' 標準プロシージャ
'----------------------------------------------
Public Const VBEXT_PK_PROC As Long = 0
