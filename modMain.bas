Attribute VB_Name = "modMain"
Option Explicit

Public VBInstance As VBIDE.VBE
Public winToolTip As VBIDE.Window
Public docToolTip As Object

Public Const APP_CATEGORY = "Microsoft Visual Basic AddIns"

Public Function InRunMode(VBInst As VBIDE.VBE) As Boolean
  InRunMode = (VBInst.CommandBars("File").Controls(1).Enabled = False)
End Function

