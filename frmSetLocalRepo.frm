VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetLocalRepo 
   Caption         =   "UserForm1"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8715
   OleObjectBlob   =   "frmSetLocalRepo.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "frmSetLocalRepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub UserForm_Initialize()
    sbLoadStrings ' Load form strings from string table.
End Sub

Private Sub sbLoadStrings()
    frmSetLocalRepo.Caption = FORM_CAPTION
    btnOK.Caption = FORM_BUTTON_OPEN_FOLDER_CAPTION
    btnExit.Caption = FORM_BUTTON_EXIT_CAPTION
End Sub

Private Sub btnOK_Click()
    Const CALLER = "frmSetLocalRepo.btnOK_Click"
    On Error GoTo btnOK_Click_ErrorHandler
    
    Dim oOpenRepo As FileDialog
    Dim vOpenFile As Variant
    Dim vFileName As Variant
    
    ' �t�@�C���_�C�A���O�I�u�W�F�N�g���Z�b�g
    '   msoFileDialogFilePicker     ���[�U�[���t�@�C����I���ł��܂��
    '   msoFileDialogFolderPicker   ���[�U�[���t�H���_��I���ł��܂��
    '   msoFileDialogOpen           ���[�U�[���t�@�C�����J�����Ƃ��ł��܂��
    '   msoFileDialogSaveAs         ���[�U�[���t�@�C����ۑ��ł��܂��
    Set oOpenRepo = Application.FileDialog(msoFileDialogFolderPicker)
    ' 1�̂ݑI��
    oOpenRepo.AllowMultiSelect = False
    ' �{�^���̕\����
    oOpenRepo.ButtonName = FILE_DIALOG_BUTTON_OPEN_CAPTION
    ' �����t�H���_���w��
    ' oOpenRepo.InitialFileName = FILE_DIALOG_DEFAULT_FOLDER
    ' �A�C�R���̑傫�����w��
    '       msoFileDialogViewDetails
    '       msoFileDialogViewLargeIcons
    '       msoFileDialogViewList
    '       msoFileDialogViewPreview
    '       msoFileDialogViewProperties
    '       msoFileDialogViewSmallIcons
    oOpenRepo.InitialView = msoFileDialogViewLargeIcons
    If (oOpenRepo.Show = -1) Then   ' �L���ȃ{�^�����N���b�N���ꂽ
        ' �I����1�ȊO�̓G���[
        If (oOpenRepo.SelectedItems.Count <> 1) Then
            MsgBox ERROR_FILE_DIALOG_MULTI_SELECT_PROMPT, vbCritical, ERROR_FILE_DIALOG_MULTI_SELECT_TITLE
        Else
            txtFolderPath.Value = oOpenRepo.SelectedItems(1)
        End If
    End If
    
    GoTo btnOK_Click_End
    
btnOK_Click_ErrorHandler:
    MsgBox "ERROR " & Err.Number & ": " & Err.Description, vbCritical, CALLER
btnOK_Click_End:
    Set oOpenRepo = Nothing
End Sub

Private Sub btnExit_Click()
    frmSetLocalRepo.Hide
End Sub

Private Sub btnExecute_Click()
    Dim oExportComponents As clsExportVBA
    
    If (txtFolderPath.Value = "") Then
        MsgBox ERROR_FORM_EMPTY_FOLDER_PROMPT, vbCritical, ERROR_FORM_EMPTY_FOLDER_TITLE
        Exit Sub
    End If
    Set oExportComponents = New clsExportVBA
    
    oExportComponents.Init (txtFolderPath.Value)
    oExportComponents.sbExportVBA
    
    txtFolderPath.Value = ""
    
    Set oExportComponents = Nothing
End Sub
