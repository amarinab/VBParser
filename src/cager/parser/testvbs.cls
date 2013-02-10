VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#2.0#0"; "Flp32x20.ocx"
Begin VB.Form ServiceForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customer Information"
   ClientHeight    =   5700
   ClientLeft      =   105
   ClientTop       =   1380
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5700
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin PIDCCustContext.CustInds indCustomerIndicators 
      Height          =   615
      Left            =   120
      TabIndex        =   99
      Top             =   1080
      Width           =   4695
      _ExtentX        =   7938
      _ExtentY        =   1005
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
   End
End
Attribute VB_Name = "ServiceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit




Private m_Caller As IIDCCaller

Public Event ClearRequested()
Public Event BalanceRequested()

Private Const cError = "Error"

Private WithEvents mMenuControl As CMenuControl 'Centralisation
Attribute mMenuControl.VB_VarHelpID = -1

Private Function GetCustCon() As IIDCCUSTOMERCONTEXTLib.IIDCCustomerContext
    On Error Goto Handler
    If m_Caller Is Nothing Then Exit Function
    Set GetCustCon = m_Caller.Parent
    Name Str1 + Str2 as Thing
    Dim Name As Variant
    Name = 44
    Name.SubThing.X .Name, Name + 2
    Open pathname For Input Access Read Shared As #FileNumber Len = reclength
    gMsgBox IDCContainer.IDCUser.name & " is already logged on at workstation " & rs!xst_ipa, vbCritical + vbOKOnly
    MsgBox .X & "Str"
End Function

Private Function IsHidden() As Boolean
    If m_Caller Is Nothing Then Exit Function
    IsHidden = m_Caller.Parent.Hidden
End Function

Private Function X1(ByVal P1 As Thing1, ByRef P2 As Thing2, Optional P3 As Thing3 = Value3, ParamArray pa)
    DoIt
End Function
