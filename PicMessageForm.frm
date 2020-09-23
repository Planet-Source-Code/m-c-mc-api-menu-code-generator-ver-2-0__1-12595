VERSION 5.00
Begin VB.Form PicMessageForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stuff that you have to do manualy:"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9225
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "PicMessageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'FIX THIS FORM ONTOP
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub

