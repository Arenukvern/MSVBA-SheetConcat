VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRuleToast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum enmControlType
  [_first]
    ectSuccess
    ectDanger
  [_last]
End Enum

Public Sub OpenToast(ControlType As enmControlType, Optional ByVal strText As String = "")
    Select Case ControlType
        Case enmControlType.ectDanger
            fDangerToast.LabelError.Caption = strText
            fDangerToast.Show
        Case enmControlType.ectSuccess
            fToast.Show
    End Select
End Sub



