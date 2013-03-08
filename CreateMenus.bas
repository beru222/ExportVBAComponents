Attribute VB_Name = "CreateMenus"
Option Explicit

'*******************************************************************
'Sub/Func   : sbCreateMenuItems
'Author     : James Rivera
'Purpose    : Create the menu item entry points for this tool.
'Arguments  : None.
'Returns    : None.
'Comments   : Uses menu item strings from StringTable.bas
'*******************************************************************
'History    : 04/02/2003 (James Rivera) - Created.
'           : 12/18/2012 (Kenichi Koyama) - modified
'*******************************************************************
Public Sub sbCreateMenuItems()

    Const CALLER = "sbCreateMenuItems"
    On Error GoTo sbCreateMenuItems_ErrorHandler

    Dim cbCtl   As Office.CommandBarControl
    Dim cbBtn   As Office.CommandBarButton
    Dim cbPopup As Office.CommandBarPopup
    
    ' First delete any old ones if they exist.
    sbRemoveMenuItems
    
    With Application.CommandBars(1).Controls
    
        ' Create top-level "Comment Tool" menu item.
        Set cbCtl = .Add(msoControlPopup, , , .Count + 1, False)
        cbCtl.Caption = MENU_EXPORTTOOLS
        cbCtl.Tag = MENU_EXPORTTOOLS
        cbCtl.BeginGroup = True
        
        ' Get a handle to the control for top-level menu item.
        Set cbPopup = cbCtl
        
        ' Create submenu for Comments creation dialog.
        Set cbBtn = cbPopup.Controls.Add(msoControlButton, , , cbPopup.Controls.Count + 1, False)
        With cbBtn
            .OnAction = "ExportVBA.sbShowForm"
            .Caption = MENU_EXPORTTOOLS_VBACOMPONENTS
            .Tag = MENU_EXPORTTOOLS_VBACOMPONENTS
        End With
    
    End With
    
    Exit Sub
sbCreateMenuItems_ErrorHandler:
    MsgBox "ERROR " & Hex(Err.Number) & ": " & Err.Description, vbCritical, CALLER
    Exit Sub
End Sub 'sbCreateMenuItems

'*******************************************************************
'Sub/Func   : sbRemoveMenuItems
'Author     : James Rivera
'Purpose    : Remove the menu item entry points for this tool.
'Arguments  : None.
'Returns    : None.
'Comments   : Uses menu item strings from StringTable.bas
'*******************************************************************
'History    : 04/02/2003 (James Rivera) - Created.
'           : 12/18/2012 (Kenichi Koyama) - modified
'*******************************************************************
Public Sub sbRemoveMenuItems()

    Const CALLER = "sbRemoveMenuItems"
    On Error GoTo sbRemoveMenuItems_ErrorHandler

    Dim cbCtl   As Office.CommandBarControl
    
    ' Delete any instances of the top-level menu if they exist.
    Set cbCtl = Application.CommandBars.FindControl(msoControlPopup, , MENU_EXPORTTOOLS)
    While Not (cbCtl Is Nothing)
        cbCtl.Delete
        Set cbCtl = Application.CommandBars.FindControl(msoControlPopup, , MENU_EXPORTTOOLS)
    Wend
    
    Exit Sub
sbRemoveMenuItems_ErrorHandler:
    MsgBox "ERROR " & Hex(Err.Number) & ": " & Err.Description, vbCritical, CALLER
    Resume Next
End Sub 'sbRemoveMenuItems

