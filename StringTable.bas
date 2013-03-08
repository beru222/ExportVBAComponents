Attribute VB_Name = "StringTable"
Option Explicit

' Menu Items
Public Const MENU_EXPORTTOOLS = "&Export Tool" ' Top level menu
Public Const MENU_EXPORTTOOLS_VBACOMPONENTS = "Export V&BAComponents" ' Calls sbShowForm()

' Form text
Public Const FORM_CAPTION = "ローカルリポジトリを選択する"
Public Const FORM_BUTTON_OPEN_FOLDER_CAPTION = "フォルダを開く"
Public Const FORM_BUTTON_EXIT_CAPTION = "終了"
Public Const FILE_DIALOG_BUTTON_OPEN_CAPTION As String = "フォルダを選択"
Public Const FILE_DIALOG_DEFAULT_FOLDER As String = ""

' Error Message
Public Const ERROR_FILE_DIALOG_MULTI_SELECT_PROMPT As String = "複数のフォルダを選択しています。"
Public Const ERROR_FILE_DIALOG_MULTI_SELECT_TITLE As String = "エラー！"
Public Const ERROR_FORM_EMPTY_FOLDER_PROMPT As String = "フォルダが選択されていません。"
Public Const ERROR_FORM_EMPTY_FOLDER_TITLE As String = "エラー！"
