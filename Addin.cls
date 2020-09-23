VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Tooltip Helper"
Option Explicit

Implements IDTExtensibility

Public WithEvents PrjHandler  As VBIDE.VBProjectsEvents          'projects event handler
Attribute PrjHandler.VB_VarHelpID = -1
Public WithEvents CmpHandler  As VBIDE.VBComponentsEvents        'components event handler
Attribute CmpHandler.VB_VarHelpID = -1
Public WithEvents CtlHandler  As VBIDE.VBControlsEvents          'controls event handler
Attribute CtlHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler As VBIDE.CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Dim mcbMenuCommandBar         As Office.CommandBarControl  'command bar object

Const guidToolTips = "{5FA32A17-DC1C-11d3-BE0C-583E06C10000}"

Public Sub Show()
On Error Resume Next
winToolTip.Visible = True
End Sub

Private Sub CmpHandler_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
If winToolTip.Visible = True Then docToolTip.Files_Changed VBComponent, 2
End Sub

Private Sub CmpHandler_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
'
End Sub

Private Sub CmpHandler_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
'
End Sub

Private Sub CmpHandler_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
If winToolTip.Visible = True Then docToolTip.Files_Changed VBComponent, 2
End Sub

Private Sub CmpHandler_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
'
End Sub

Private Sub CmpHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
If winToolTip.Visible = True Then docToolTip.Files_Changed VBComponent, 2
End Sub

Private Sub CtlHandler_ItemAdded(ByVal VBControl As VBIDE.VBControl)
If winToolTip.Visible = True Then docToolTip.Controls_Changed VBControl, 1
End Sub

Private Sub CtlHandler_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
If winToolTip.Visible = True Then docToolTip.Controls_Changed VBControl, 2
End Sub

Private Sub CtlHandler_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
If winToolTip.Visible = True Then docToolTip.Controls_Changed VBControl, 3
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    
Dim aiTmp As VBIDE.AddIn

On Error GoTo error_handler
    
'save the vb instance
Set VBInstance = VBInst

If Not winToolTip Is Nothing Then
  'already running so just show it
  Show
  If ConnectMode = vbext_cm_AfterStartup Then
    'started from the addin manager
    AddToCommandBar
  End If
  Exit Sub
End If
    
'create the tool window
If ConnectMode = vbext_cm_External Then
  'need to see if it is already running
  On Error Resume Next
  Set aiTmp = VBInstance.Addins("ToolTips.Connect")
  On Error GoTo error_handler
  If aiTmp Is Nothing Then
    'app is not in the VBADDIN.INI file so it is not in the collection
    'so lets attempt to use the 1st addin in the collection just
    'to get this app running and if there are none, an error
    'will occur and this app will not run
    Set winToolTip = VBInstance.Windows.CreateToolWindow(VBInstance.Addins(1), "ToolTip.udocToolTips", "ToolTips Helper", guidToolTips, docToolTip)
  Else
    If aiTmp.Connect = False Then
      Set winToolTip = VBInstance.Windows.CreateToolWindow(aiTmp, "ToolTip.udocToolTips", "ToolTips Helper", guidToolTips, docToolTip)
    End If
  End If
Else
  'must've been called from addin mgr
  Set winToolTip = VBInstance.Windows.CreateToolWindow(AddInInst, "ToolTip.udocToolTips", "ToolTips Helper", guidToolTips, docToolTip)
End If
    
'sink the project, components and controls event handler
Set Me.PrjHandler = VBInstance.Events.VBProjectsEvents
Set Me.CmpHandler = VBInstance.Events.VBComponentsEvents(Nothing)
Set Me.CtlHandler = VBInstance.Events.VBControlsEvents(Nothing, Nothing)

If ConnectMode = vbext_cm_External Then
  'started from the addin toolbar
  Show
ElseIf ConnectMode = vbext_cm_AfterStartup Then
  'started from the addin manager
  AddToCommandBar
End If

Exit Sub
    
error_handler:
MsgBox Err.Description

End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
    
On Error Resume Next
    
'delete the command bar entry
mcbMenuCommandBar.Delete
Set winToolTip = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
'    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
'        'set this to display the form on connect
'        Me.Show
'    End If
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
Show
End Sub

Public Sub AddToCommandBar()

On Error GoTo errhand

'make sure the standard toolbar is visible
VBInstance.CommandBars(2).Visible = True

'add it to the command bar
'the following line will add the ToolTips manager to the
'Standard toolbar to the right of the ToolBox button
Set mcbMenuCommandBar = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
'set the caption
mcbMenuCommandBar.Caption = "ToolTips Helper"
'copy the icon to the clipboard
Clipboard.SetData LoadResPicture(101, 0)
'set the icon for the button
mcbMenuCommandBar.PasteFace

'sink the event
Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)

Exit Sub

errhand:
MsgBox Err.Description & ", " & Err.Source

End Sub

Private Sub PrjHandler_ItemActivated(ByVal VBProject As VBIDE.VBProject)
' No need to update anything
End Sub

Private Sub PrjHandler_ItemAdded(ByVal VBProject As VBIDE.VBProject)
' No need to update anything
End Sub

Private Sub PrjHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
' Take care of the list if the current project is removed
If winToolTip.Visible = True Then docToolTip.Project_Changed VBProject, 1
End Sub

Private Sub PrjHandler_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
' No need to update anything
End Sub
