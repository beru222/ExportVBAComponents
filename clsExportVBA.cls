VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExportVBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cFilePath As Variant    ' 保存先パス

Private Sub Class_initialize()
    cFilePath = Empty
End Sub

Private Sub Class_Terminate()
    cFilePath = Empty
End Sub

'*******************************************************************
'Sub/Func   : sbExportVBA
'Author     : ANET Corporation.
'Purpose    :
'Arguments  : none.
'Returns    : none.
'Comments   :
'*******************************************************************
'History    : 2013/02/18 (Kenichi Koyama) - Created.
'*******************************************************************
Public Sub sbExportVBA()
    Const CALLER = "clsExportVBA.sbExportVBA"
    On Error GoTo sbExportVBA_ErrorHandler
    
    Dim oVBComponent As VBComponent
    Dim vFileExtension As Variant
    Dim vFileName As Variant

    For Each oVBComponent In ActiveWorkbook.VBProject.VBComponents
        ' 適切な拡張子を付ける
        vFileExtension = fnGetExtFromModuleType(oVBComponent.Type)
        If (Not IsEmpty(vFileExtension)) Then
            vFileName = cFilePath _
                            & "\" _
                            & oVBComponent.Name _
                            & "." _
                            & vFileExtension
            Call oVBComponent.Export(vFileName)
#If (DEBUG_FLAG = 1) Then
            Debug.Print "Save " & oVBComponent.Name
#End If
            vFileName = Empty
            vFileExtension = Empty
        End If
    Next oVBComponent
    
    GoTo sbExportVBA_End
    
sbExportVBA_ErrorHandler:
    MsgBox "ERROR " & Err.Number & ": " & Err.Description, vbCritical, CALLER
sbExportVBA_End:
    Set oVBComponent = Nothing
End Sub 'sbExportVBA

Private Function fnGetExtFromModuleType(aType As Integer) As Variant
    fnGetExtFromModuleType = Empty
    ' 指定されたモジュール・タイプに対応する拡張子を返す
    Select Case aType
        Case vbext_ct_StdModule
            fnGetExtFromModuleType = "bas"
        Case vbext_ct_ClassModule, vbext_ct_Document
            fnGetExtFromModuleType = "cls"
        Case vbext_ct_MSForm
            fnGetExtFromModuleType = "frm"
    End Select
End Function

'*******************************************************************
'Sub/Func   : Init
'Author     : ANET Corporation.
'Purpose    :
'Arguments  : string of file path.
'Returns    :
'Comments   :
'*******************************************************************
'History    : 2013/02/18 (Kenichi Koyama) - Created.
'*******************************************************************
Public Sub Init( _
            ByVal vFilePath As Variant)
    Const CALLER = "clsExportVBA.Init"
    On Error GoTo Init_ErrorHandler
    
    cFilePath = vFilePath
    
    Exit Sub
Init_ErrorHandler:
    MsgBox "ERROR " & Err.Number & ": " & Err.Description, vbCritical, CALLER
End Sub 'Init

