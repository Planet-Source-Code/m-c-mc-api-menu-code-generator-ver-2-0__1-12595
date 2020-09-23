VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "MC API Menu Code Generator ver 2.0 (Pro)"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1560
      Top             =   5280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Help me "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   35
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000012&
      Caption         =   "General menu behaviour settings"
      ForeColor       =   &H000000FF&
      Height          =   5175
      Left            =   1680
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton Command7 
         Caption         =   "OK"
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   2880
         Width           =   2535
      End
      Begin VB.PictureBox Picture4 
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   6915
         TabIndex        =   17
         Top             =   1560
         Width           =   6975
         Begin VB.OptionButton Option7 
            Caption         =   "PopUpMenu will be triggered with left click."
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Value           =   -1  'True
            Width           =   5775
         End
         Begin VB.OptionButton Option7 
            Caption         =   "PopUpMenu will be triggered with right click."
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   5775
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   855
            Left            =   4080
            TabIndex        =   36
            Top             =   -240
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   6915
         TabIndex        =   13
         Top             =   360
         Width           =   6975
         Begin VB.OptionButton Option6 
            Caption         =   "PopUpMenu will appear to the center/down side where mouse will be clicked. "
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   16
            Top             =   720
            Width           =   5895
         End
         Begin VB.OptionButton Option6 
            Caption         =   "PopUpMenu will appear to the right/down side where mouse will be clicked. "
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   360
            Width           =   5895
         End
         Begin VB.OptionButton Option6 
            Caption         =   "PopUpMenu will appear to the left/down side where mouse will be clicked. "
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Value           =   -1  'True
            Width           =   5895
         End
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Know'n bugs "
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   3600
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   2475
      TabIndex        =   29
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
      Begin VB.PictureBox MenuPicContainer 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   3
         Left            =   1200
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   33
         Top             =   120
         Width           =   240
      End
      Begin VB.PictureBox MenuPicContainer 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "Form1.frx":01F2
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   32
         Top             =   120
         Width           =   240
      End
      Begin VB.PictureBox MenuPicContainer 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "Form1.frx":03E4
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   31
         Top             =   120
         Width           =   240
      End
      Begin VB.PictureBox MenuPicContainer 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "Form1.frx":05D6
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   30
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Beginner Help"
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00808080&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   23
      Top             =   1440
      Width           =   1695
      Begin VB.CommandButton Command10 
         Caption         =   "Generate Code"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Paste code direct under event where you want POPUPmenu to appear, i.e. Command1_Click. Must have module with pub declares."
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Public Declares .."
         Height          =   255
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Declarations ready to be pasted into module.Once u have mod. u yust need 'Code No Declares' button."
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Private declares"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Paste this into declaration section of new form. "
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Copy code to CLP"
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ver 2.0 info."
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   3240
      Width           =   1575
   End
   Begin VB.ListBox List4 
      Height          =   4935
      Left            =   8520
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.ListBox List3 
      Height          =   4935
      ItemData        =   "Form1.frx":07C8
      Left            =   7560
      List            =   "Form1.frx":07CA
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   1005
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   5520
      Width           =   8295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load structure"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "map"
      Filter          =   "Structures (*.MAP)|*.MAP;"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save structure for later use"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Exit"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   4935
      ItemData        =   "Form1.frx":07CC
      Left            =   6600
      List            =   "Form1.frx":07CE
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   4935
      ItemData        =   "Form1.frx":07D0
      Left            =   1680
      List            =   "Form1.frx":07D2
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click here to vote for !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6000
      TabIndex        =   27
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Pic"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   8520
      TabIndex        =   11
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Item state"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   9
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Item type"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   4
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Structure building box"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timervar As Boolean 'under timer 1
Dim controler
Dim ClickNoCanDo As Boolean 'some click trouble repair
Dim VariableIndex As Integer 'for checked items needed variables
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Function GetWindowsDirectory() As String
   Dim s As String
   Dim i As Integer
   i = GetWindowsDirectoryA("", 0)
   s = Space(i)
   Call GetWindowsDirectoryA(s, i)
   GetWindowsDirectory = AddBackslash(Left$(s, i - 1))
End Function
Public Function AddBackslash(s As String) As String
   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s + "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If
End Function
'above is not esential for this project, allso not mycode
Private Sub PrivateDeclaresForCode()


Text2.SetFocus
Text2.Text = ""

'1.declaration section
SendKeys "'Declaration section"
SendKeys "{ENTER}"
SendKeys "Private Declare Function CreatePopupMenu Lib" & Chr(34) & "user32.dll" & Chr(34) & " {(}{)}  As Long"
SendKeys "{ENTER}"
SendKeys "Private Declare Function DestroyMenu Lib " & Chr(34) & "user32.dll" & Chr(34) & " {(}ByVal hMenu As Long{)} As Long"
SendKeys "{ENTER}"
SendKeys "Private Type MENUITEMINFO"
SendKeys "{ENTER}"
SendKeys "        cbSize As Long"
SendKeys "{ENTER}"
SendKeys "        fMask As Long"
SendKeys "{ENTER}"
SendKeys "        fType As Long"
SendKeys "{ENTER}"
SendKeys "        fState As Long"
SendKeys "{ENTER}"
SendKeys "        wID As Long"
SendKeys "{ENTER}"
SendKeys "        hSubMenu As Long"
SendKeys "{ENTER}"
SendKeys "        hbmpChecked As Long"
SendKeys "{ENTER}"
SendKeys "        hbmpUnchecked As Long"
SendKeys "{ENTER}"
SendKeys "        dwItemData As Long"
SendKeys "{ENTER}"
SendKeys "        dwTypeData As String"
SendKeys "{ENTER}"
SendKeys "        cch As Long"
SendKeys "{ENTER}"
SendKeys "End Type"
SendKeys "{ENTER}"

'Constant Definitions
 
SendKeys "Private Const MIIM_STATE = &H1"
SendKeys "{ENTER}"
SendKeys "Private Const MIIM_ID = &H2"
SendKeys "{ENTER}"
SendKeys "Private Const MIIM_SUBMENU = &H4"
SendKeys "{ENTER}"
SendKeys "Private Const MIIM_CHECKMARKS = &H8"
SendKeys "{ENTER}"
SendKeys "Private Const MIIM_DATA = &H20"
SendKeys "{ENTER}"
SendKeys "Private Const MIIM_TYPE = &H10"
SendKeys "{ENTER}"
SendKeys "Private Const MFT_BITMAP = &H4"
SendKeys "{ENTER}"
SendKeys "Private Const MFT_MENUBARBREAK = &H20"
SendKeys "{ENTER}"
SendKeys "Private Const MFT_MENUBREAK = &H40"
SendKeys "{ENTER}"
SendKeys "Private Const MFT_OWNERDRAW = &H100"
SendKeys "{ENTER}"
SendKeys "Private Const MFT_RADIOCHECK = &H200"
SendKeys "{ENTER}"
SendKeys "Private Const MFT_RIGHTJUSTIFY = &H4000"
SendKeys "{ENTER}"
SendKeys "Private Const MFT_RIGHTORDER = &H2000"
SendKeys "{ENTER}"
SendKeys "Private Const MFT_SEPARATOR = &H800"
SendKeys "{ENTER}"
SendKeys "Private Const MFT_STRING = &H0"
SendKeys "{ENTER}"
SendKeys "Private Const MFS_CHECKED = &H8"
SendKeys "{ENTER}"
SendKeys "Private Const MFS_DEFAULT = &H1000"
SendKeys "{ENTER}"
SendKeys "Private Const MFS_DISABLED = &H2"
SendKeys "{ENTER}"
SendKeys "Private Const MFS_ENABLED = &H0"
SendKeys "{ENTER}"
SendKeys "Private Const MFS_GRAYED = &H1"
SendKeys "{ENTER}"
SendKeys "Private Const MFS_HILITE = &H80"
SendKeys "{ENTER}"
SendKeys "Private Const MFS_UNCHECKED = &H0"
SendKeys "{ENTER}"
SendKeys "Private Const MFS_UNHILITE = &H0"

'functions = API-s
SendKeys "{ENTER}"
SendKeys "Private Declare Function InsertMenuItem Lib " & Chr(34) & "user32.dll" & Chr(34) & " Alias " & Chr(34) & "InsertMenuItemA" & Chr(34) & " _"
SendKeys "{ENTER}"
SendKeys "{(}ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO{)} As Long"
SendKeys "{ENTER}"
SendKeys "Private Declare Function TrackPopupMenu Lib " & Chr(34) & "user32.dll" & Chr(34) & " _"
SendKeys "{ENTER}"
SendKeys "{(}ByVal hMenu As Long, ByVal uFlags As Long, ByVal x As Long, ByVal y As Long," & " _"
SendKeys "{ENTER}"
SendKeys "ByVal nReserved As Long, ByVal hWnd As Long, ByVal prcRect As Long{)} As Long"
SendKeys "{ENTER}"
SendKeys "Private Const TPM_RIGHTALIGN = &H8&"
SendKeys "{ENTER}"
SendKeys "Private Const TPM_CENTERALIGN = &H4&"
SendKeys "{ENTER}"
SendKeys "Private Const TPM_LEFTALIGN = &H0"
SendKeys "{ENTER}"
SendKeys "Private Const TPM_TOPALIGN = &H0"
SendKeys "{ENTER}"
SendKeys "Private Const TPM_NONOTIFY = &H80"
SendKeys "{ENTER}"
SendKeys "Private Const TPM_RETURNCMD = &H100"
SendKeys "{ENTER}"
SendKeys "Private Const TPM_LEFTBUTTON = &H0"
SendKeys "{ENTER}"
SendKeys "Private Const  TPM_RIGHTBUTTON = &H2&"
SendKeys "{ENTER}"
SendKeys "Private Type POINT_TYPE"
SendKeys "{ENTER}"
SendKeys "x As Long"
SendKeys "{ENTER}"
SendKeys "y As Long"
SendKeys "{ENTER}"
SendKeys "End Type"
SendKeys "{ENTER}"
SendKeys "Private Declare Function GetCursorPos Lib " & Chr(34) & "user32.dll" & Chr(34) & " {(}lpPoint As POINT_TYPE{)} As Long"
SendKeys "{ENTER}"
SendKeys "Private Declare Function AppendMenu Lib " & Chr(34) & "user32" & Chr(34) & " Alias " & Chr(34) & "AppendMenuA" & Chr(34) & " {(}ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any{)} As Long"
SendKeys "{ENTER}"
SendKeys "Private Declare Function SetMenuItemBitmaps Lib " & Chr(34) & "user32" & Chr(34) & " {(}ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long{)} As Long"
SendKeys "{ENTER}"
End Sub

Private Function GetItemLevel(ItemCaption As String)
If Left(ItemCaption, 1) <> "-" Then
GetItemLevel = 0
Else:
    For i = 0 To Len(ItemCaption)
           If Mid(ItemCaption, i + 1, 1) <> "-" Then GetItemLevel = Counter / 4: Exit For
           Counter = Counter + 1
    Next i
End If
End Function
Private Function GetItemCaption(ItemCaption As String)
If Left(ItemCaption, 1) <> "-" Then
GetItemCaption = ItemCaption
Else:
    For i = 0 To Len(ItemCaption)

           If Mid(ItemCaption, i + 1, 1) <> "-" Then GetItemCaption = Right(ItemCaption, Len(ItemCaption) - Counter): Exit For
           Counter = Counter + 1
    Next i
End If

End Function
Private Sub CheckedItemProblemProcedure(InitialState As String, MenuNum As Integer, CompletePath As String, ItemType As String, ItemMagicNum As String, Caption As String, ItemSubMenu As Integer, ItemState As String)
 SendKeys "{ENTER}"
 SendKeys "{ENTER}"
 If InitialState = "MFS_CHECKED" Then
 SendKeys "Case False"
 Else
 SendKeys "Case True"
 End If
 
      SendKeys "{ENTER}"
      SendKeys "With mii" & MenuNum & "'{(}" & CompletePath & "{)}"
      SendKeys "{ENTER}"
            'type thing
            SendKeys ".fType =" & ItemType
      SendKeys "{ENTER}"
            'state thing
             If InitialState = "MFS_CHECKED" Then
             SendKeys ".fState = " & Left(ItemState, Len(ItemState) - Len(InitialState)) & "MFS_UNCHECKED"
             Else
             SendKeys ".fState = " & Left(ItemState, Len(ItemState) - Len(InitialState)) & "MFS_CHECKED"
             End If
      SendKeys "{ENTER}"
      SendKeys ".wID =" & ItemMagicNum & "' Assign this item an item identifier."
      SendKeys "{ENTER}"
      SendKeys ".dwTypeData =" & Chr(34) & Caption & Chr(34) ' Display the following text for the item."
      SendKeys "{ENTER}"
      SendKeys ".cch = Len{(}" & Chr(34) & Caption & Chr(34) & "{)}"
      SendKeys "{ENTER}"
      'if there is a submenu for this item then......
      If ItemSubMenu > 0 Then
      SendKeys ".hSubMenu = hPopupMenu" & ItemSubMenu
      SendKeys "{ENTER}"
      Else
      SendKeys ".hSubMenu = 0"
      SendKeys "{ENTER}"
      End If
      SendKeys "End With"
      
SendKeys "{ENTER}"
SendKeys "{ENTER}"
SendKeys "Case Else"
SendKeys "{ENTER}"
SendKeys "End Select"
SendKeys "{ENTER}"
SendKeys "'--------------------------End of complicating about checked/unchecked-----------------------------"
SendKeys "{ENTER}"
End Sub
Private Sub Command1_Click()
Text2.SetFocus
Text2.Text = ""
PrivateDeclaresForCode
End Sub

Private Sub ApiPopUpMenuCodeGenerator() 'code maker
'avoid disaster
If List1.ListCount = 0 Then
MsgBox " No items in list box, ending...."
Exit Sub
End If

Dim magicnumber As Integer

magicnumber = 1000 ' start value of it, could be any other that comes to your mind"
MenuAppendNum = 1


'SAVE STUFF INTO TABLE --------------------------------------------------------------------

'dbs stuff---------------------
Dim dbs As Database
Dim rst As Recordset
Set dbs = OpenDatabase(App.Path & "\my.mdb")
Set rst = dbs.OpenRecordset("SaveConstruct")
'------------------------------
' Clear table.
dbs.Execute "DELETE * FROM SaveConstruct ;"

For i = 0 To List1.ListCount - 1
With rst


.AddNew
'ControlIndex
![ControlIndex] = i
'items string
![Container] = List1.List(i)
'level
![Level] = GetItemLevel(List1.List(i))
'caption = cut off ---- signs from string if any
![Caption] = GetItemCaption(List1.List(i))

'-----inserting this in version 2.00----TRANSLASTION-----------------------------
Dim TypeOutputString As String
Dim StateOutputString As String

'1.first translate state
'Type translated......
        TypeOutputString = "MFT_STRING"
        
        Select Case Mid(List2.List(i), 2, 1)
        Case "R"
        TypeOutputString = TypeOutputString & " Or MFT_RADIOCHECK"
        Case "C"
        'do nothing
        Case Else
        End Select

        Select Case Mid(List2.List(i), 3, 1)
        Case "L"
        TypeOutputString = TypeOutputString & " Or MFT_MENUBARBREAK"
        Case "N"
        TypeOutputString = TypeOutputString & " Or MFT_MENUBREAK"
        Case Else
        End Select
'2.second translate state
'State translated......
        Select Case Mid(List3.List(i), 1, 1)
        Case "E"
        StateOutputString = "MFS_ENABLED"
        Case "D"
        StateOutputString = "MFS_DISABLED"
        Case "G"
        StateOutputString = "MFS_GRAYED"
        Case Else
        End Select
        
        Select Case Mid(List3.List(i), 2, 1)
        Case "B"
        StateOutputString = StateOutputString & " Or MFS_DEFAULT"
        Case "N"
        'do nothing
        Case Else
        End Select
        
        Select Case Mid(List3.List(i), 3, 1)
        Case "C"
        StateOutputString = StateOutputString & " Or MFS_CHECKED"
        ![Checked] = "MFS_CHECKED"
        Case "U"
        StateOutputString = StateOutputString & " Or MFS_UNCHECKED"
        ![Checked] = "MFS_UNCHECKED"
        Case Else
        End Select
        


'-------------------END OF TRANSLATION---------------------------------------------------

'type of
        If GetItemCaption(List1.List(i)) = "/separator/" Then
        ![ItemType] = "MFT_SEPARATOR"
        Else
        ![ItemType] = TypeOutputString
        End If
'State
![ItemState] = StateOutputString

'Picture Yes/no
![Picture] = List4.List(i)

'Place = order in which items in menu should appear
        b = 0
        c = 0
        a = GetItemLevel(List1.List(i))
        Do
        If i = 0 Then ![Place] = 0: Exit Do
        b = b + 1
        If i - b < 0 Then ![Place] = c: Exit Do

                Select Case GetItemLevel(List1.List(i - b))
                Case Is = a
                c = c + 1
                Case Is < a
                ![Place] = c
                c = 0
                Exit Do
                Case Is > a
            
                Case Else
                Beep '? - if this happen then throw your computer away, lol
                End Select
        Loop Until i - b < 0
.Update
End With
Next i
'END SAVING  STUFF INTO TABLE


DoEvents
'dbs stuff---------------------
Dim rst1 As Recordset
Set rst1 = dbs.OpenRecordset("SaveConstruct")

    
'find max level in list1
MaxItemLevel = 0
For i = 0 To List1.ListCount
        If GetItemLevel(List1.List(i)) > MaxItemLevel Then
        MaxItemLevel = GetItemLevel(List1.List(i))
        End If
Next i


s = MaxItemLevel
'now read from bottom to top & get data into table
For j = 0 To s
       rst1.MoveLast
       d = rst1![Caption]
       Do
                     If rst1![Level] = MaxItemLevel Then
                        Do
                           If rst1![Level] = MaxItemLevel Then
                                With rst1
                                .Edit
                                 ![MenuNum] = MenuAppendNum 'which menu item fits in
                                 ![MenuName] = "hPopupMenu" & MenuAppendNum 'which menu item fits in
                                 ![ItemMagicNum] = magicnumber 'allso for separator
                                .Update
                                End With
                          magicnumber = magicnumber + 1
                           End If
                           rst1.MovePrevious
                        If rst1.BOF Then Exit For
                        Loop Until rst1![Level] < MaxItemLevel
                   rst1.MoveNext
                   MenuAppendNum = MenuAppendNum + 1
                   End If
        rst1.MovePrevious
        Loop Until rst1.BOF
MaxItemLevel = MaxItemLevel - 1
'now next level
Next j

'Enter sub menus data - which menu is to be opened when mouse moves over specific item
For i = 0 To List1.ListCount

        Select Case GetItemCaption(List1.List(i))
        Case Is = "/separator/" 'separators can't have submenus !
            
                       If GetItemLevel(List1.List(i + 1)) > GetItemLevel(List1.List(i)) Then 'problem at sight
                       t = i
                       Do
                       t = t - 1
                       Loop Until GetItemCaption(List1.List(t)) <> "/separator/"
                       'now we have menu (t) and it's submenu (i+1)... get it into database
                        
                        rst1.MoveFirst
                        Do
                                If rst1![Caption] = GetItemCaption(List1.List(i + 1)) And rst1![ControlIndex] = i + 1 Then
                                SubMen = rst1![MenuNum]
                                Exit Do
                                End If
                                rst1.MoveNext
                        Loop
                        
                        rst1.MoveFirst
                        Do
                                If rst1![Caption] = GetItemCaption(List1.List(t)) And rst1![ControlIndex] = t Then
                                     With rst1
                                     .Edit
                                     ![ItemSubMenu] = SubMen
                                     .Update
                                     End With
                                     Exit Do
                                End If
                        rst1.MoveNext
                        Loop
                
                
                Else: GoTo 11 'do nothing as 'separators can't have submenus !
                End If
        
        Case Else
        If GetItemLevel(List1.List(i + 1)) = (GetItemLevel(List1.List(i)) + 1) Then
        'if next one level higher as curent item then .....
                rst1.MoveFirst
                Do
                        If rst1![Caption] = GetItemCaption(List1.List(i + 1)) And rst1![ControlIndex] = i + 1 Then
                        SubMen = rst1![MenuNum]
                        Exit Do
                        End If
                        rst1.MoveNext
                Loop
                
                rst1.MoveFirst
                Do
                        If rst1![Caption] = GetItemCaption(List1.List(i)) And rst1![ControlIndex] = i Then
                             With rst1
                             .Edit
                             ![ItemSubMenu] = SubMen
                             .Update
                             End With
                             Exit Do
                        End If
                rst1.MoveNext
                Loop
       End If
       End Select
11
Next i

'--------------------------------------------------------------------------------------------
'added into ver 2.0
'Coplete Path Finder Procedure
'Before writing code - get some data that will place output code coments on steroids
'rst1 is steel opened ....., so use it
rst1.MoveFirst
For i = 0 To List1.ListCount - 1
    depth = rst1![Level]
    Select Case depth
    Case 0
            With rst1
            .Edit
            ![CompletePath] = ![Caption]
            .Update
            End With
    Case Else
                  Dim FCP As String 'final complete path
                  FCP = GetItemCaption(List1.List(i))
                          f = i
                          CurrentLevel = GetItemLevel(List1.List(i))
                          Do
                                 f = f - 1
                                 If GetItemLevel(List1.List(f)) < CurrentLevel And _
                                 GetItemCaption(List1.List(f)) <> "/Separator/" Then
                                 FCP = GetItemCaption(List1.List(f)) & "/" & FCP
                                 End If
                                 CurrentLevel = GetItemLevel(List1.List(f))
                                 
                          Loop Until GetItemLevel(List1.List(f)) = 0 'at 0 exit loop
                                'path is complete - write it down
                                With rst1
                                .Edit
                                ![CompletePath] = FCP
                                .Update
                                End With
                                                    
    End Select
    
rst1.MoveNext
Next i

'--------------------------------------------------------------------------------------------
'START WRITING CODE
'--------------------------------------------------------------------------------------------

'only god knows why these lines are here, but they must be !
'comment made one month after ver 1 coding - learn from it, lol
If controler = 0 Then
Text2.SetFocus
Text2.Text = ""
controler = 1
End If

    'lets see how many menus is to be created
    
     NumOfMenusToBeCreated = 1
     rst1.MoveFirst
     For i = 0 To List1.ListCount - 1
     If rst1![MenuNum] > NumOfMenusToBeCreated Then NumOfMenusToBeCreated = rst1![MenuNum]
     rst1.MoveNext
     Next i



'2. The core of the thing

      SendKeys "'--------------------------------------------------------------------------------------------------------------------------"
      SendKeys "{ENTER}"
      SendKeys "'CODE AUTOGENERATED WITH:  MC API Menu Code Generator ver 2.0 "
      SendKeys "{ENTER}"
      SendKeys "'---------------------------------------------------------------------------------------------------------------------------"
      SendKeys "{ENTER}"
      
            For i = 1 To NumOfMenusToBeCreated
            SendKeys "Dim hPopupMenu" & i & " As Long' handle to the popup menu to display"
            SendKeys "{ENTER}"
            Next i
    
            For i = 1 To NumOfMenusToBeCreated
            SendKeys "Dim mii" & i & " As MENUITEMINFO   ' describes menu items to add"
            SendKeys "{ENTER}"
            Next i

        
      SendKeys "Dim curpos As POINT_TYPE  ' holds the current mouse coordinates"
      SendKeys "{ENTER}"
      SendKeys "Dim menusel As Long       ' ID of what the user selected in the popup menu"
      SendKeys "{ENTER}"
      SendKeys "Dim retval As Long        ' generic return value"
      SendKeys "{ENTER}"
      
      SendKeys "{ENTER}"
  
      SendKeys "{ENTER}"
    
   
    'Create the popup menus which are initialy empty.
    SendKeys "'Create the popup menus which are initialy empty."
    SendKeys "{ENTER}"
    For i = 1 To NumOfMenusToBeCreated
    SendKeys "hPopupMenu" & i & " = CreatePopupMenu{(}{)}"
    SendKeys "{ENTER}"
    Next i
    SendKeys "{ENTER}"
    
      'Structure of menus to be displayed
      SendKeys "'Create the structure which is the base for all menus:"
      SendKeys "{ENTER}"
      SendKeys "With mii1"
      SendKeys "{ENTER}"
      SendKeys ".cbSize = Len{(}mii1{)}' The size of this structure."
      SendKeys "{ENTER}"
      SendKeys ".fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU' Which elements of the structure to use."
      SendKeys "{ENTER}"
      SendKeys "End With"
      SendKeys "{ENTER}"
      SendKeys "{ENTER}"
      
   SendKeys "'Make all structures equal"
   SendKeys "{ENTER}"
   For i = 2 To NumOfMenusToBeCreated
            SendKeys "mii" & i & " = mii1"
            SendKeys "{ENTER}"
            Next i
      
  Dim rst2 As Recordset
  Set rst2 = dbs.OpenRecordset("SaveConstructQuery") 'go get data
  rst2.MoveFirst
  DoEvents



'get items into created menus ! & describe their properties"
'-------------------------------------------
'-------------------------------------------
'-------------------------------------------
For i = 0 To List1.ListCount - 1

 'special treatment for those items that you selected 'check option/Yes or No'
      If rst2![Checked] = "MFS_CHECKED" Or rst2![Checked] = "MFS_UNCHECKED" Then
      VariableIndex = VariableIndex + 1
      SendKeys "'--------------------------Complications about checked/unchecked-----------------------------"
      SendKeys "{ENTER}"
      SendKeys "Select Case APICheck" & VariableIndex
                    If rst2![Checked] = "MFS_CHECKED" Then
                    SendKeys "{ENTER}"
                    SendKeys "{ENTER}"
                    SendKeys "Case True"
                    Else
                    SendKeys "{ENTER}"
                    SendKeys "{ENTER}"
                    SendKeys "Case False"
                    End If
      End If
 'end of special treatment..................................................................
 
      SendKeys "{ENTER}"
      SendKeys "With mii" & rst2![MenuNum] & "'{(}" & rst2![CompletePath] & "{)}"
      SendKeys "{ENTER}"
            'type thing
            SendKeys ".fType =" & rst2![ItemType]
      SendKeys "{ENTER}"
            'state thing
            SendKeys ".fState =" & rst2![ItemState]
      SendKeys "{ENTER}"
      SendKeys ".wID =" & rst2![ItemMagicNum] & "' Assign this item an item identifier."
      SendKeys "{ENTER}"
      SendKeys ".dwTypeData =" & Chr(34) & rst2![Caption] & Chr(34) ' Display the following text for the item."
      SendKeys "{ENTER}"
      SendKeys ".cch = Len{(}" & Chr(34) & rst2![Caption] & Chr(34) & "{)}"
      SendKeys "{ENTER}"
      'if there is a submenu for this item then......
      If rst2![ItemSubMenu] > 0 Then
      SendKeys ".hSubMenu = hPopupMenu" & rst2![ItemSubMenu]
      SendKeys "{ENTER}"
      Else
      SendKeys ".hSubMenu = 0"
      SendKeys "{ENTER}"
      End If
      SendKeys "End With"
      
    'add on for those items that you selected 'check option/Yes or No'
    If rst2![Checked] = "MFS_CHECKED" Or rst2![Checked] = "MFS_UNCHECKED" Then
    CheckedItemProblemProcedure rst2![Checked], rst2![MenuNum], rst2![CompletePath], rst2![ItemType], rst2![ItemMagicNum], rst2![Caption], rst2![ItemSubMenu], rst2![ItemState]
    End If
    'end of add on for those items that you selected 'check option/Yes'
      
    '-------------------------------------------
    '-------------------------------------------
    '-------------------------------------------
      
      SendKeys "{ENTER}"
      'well where to send the item ?
      SendKeys "retval = InsertMenuItem{(}" & rst2![MenuName] & "," & rst2![Place] & ",1, mii" & rst2![MenuNum] & "{)}"
      SendKeys "{ENTER}"
      rst2.MoveNext

      
Next i

SendKeys "{ENTER}"
SendKeys "'The following code is for adding pictures into menus, if there are any!"
SendKeys "{ENTER}"
SendKeys "'------------------------------------------------------------"
SendKeys "{ENTER}"
SendKeys "'------------------------------------------------------------"
SendKeys "{ENTER}"

'HERE ADD PICTURES IF ANY
rst2.MoveFirst 'back to beginning
Dim picindex As Integer
Dim PicMsgBoxWarning As Boolean
picindex = 0
For i = 0 To List1.ListCount - 1
        If rst2![Picture] = "Yes" Then
                If Mid(List2.List(i), 2, 1) = "P" Then
                SendKeys "{ENTER}"
                SendKeys "retval = SetMenuItemBitmaps{(}" & rst2![MenuName] & "," & rst2![ItemMagicNum] & ",1, MenuPicContainer {(}" & picindex & "{)}," & "MenuPicContainer {(}" & picindex + 1 & "{)} {)}"
                picindex = picindex + 2
                Else
                SendKeys "{ENTER}"
                SendKeys "retval = SetMenuItemBitmaps{(}" & rst2![MenuName] & "," & rst2![ItemMagicNum] & ",1, MenuPicContainer {(}" & picindex & "{)}," & "MenuPicContainer {(}" & picindex & "{)} {)}"
                picindex = picindex + 1
                
                End If
                PicMsgBoxWarning = True
        End If
rst2.MoveNext
Next i

SendKeys "{ENTER}"
SendKeys "'------------------------------------------------------------"
SendKeys "{ENTER}"
SendKeys "'------------------------------------------------------------"
SendKeys "{ENTER}"
SendKeys "{ENTER}"

    ' Determine where the mouse cursor currently is, in order to have
        ' the popup menu appear at that point.
         SendKeys "retval = GetCursorPos{(}curpos{)}"
         SendKeys "{ENTER}"
 ' Display the popup menu at the mouse cursor.  Instead of sending messages
        ' to window Form1, have the function merely return the ID of the user's selection.
        Dim BehaviourString As String
                       BehaviourString = "TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD"
                          
                       If Option6(0).Value = True Then BehaviourString = BehaviourString & " Or TPM_RIGHTALIGN"
                       If Option6(1).Value = True Then BehaviourString = BehaviourString & " Or TPM_LEFTALIGN"
                       If Option6(2).Value = True Then BehaviourString = BehaviourString & " Or TPM_CENTERALIGN"

                       If Option7(0).Value = True Then BehaviourString = BehaviourString & " Or TPM_LEFTBUTTON"
                       If Option7(1).Value = True Then BehaviourString = BehaviourString & " Or TPM_RIGHTBUTTON"
                           
                          
        SendKeys "menusel = TrackPopupMenu{(}hPopupMenu" & NumOfMenusToBeCreated & "," & BehaviourString & ",curpos.x, curpos.y, 0, Form1.hWnd, 0 {)}"
        SendKeys "{ENTER}"
      ' Before acting upon the user's selection, destroy the popup menu now.
        SendKeys "retval = DestroyMenu{(}hPopupMenu" & NumOfMenusToBeCreated & "{)}"
        SendKeys "{ENTER}"
        
        
        
        'create event handling
        SendKeys "'------------------------------------------------------------------------------------------------"
        SendKeys "{ENTER}"
        SendKeys "'DOWN BELOW  PUT IN YOUR CODE MANUALY !!!!"
        SendKeys "{ENTER}"
        SendKeys "'------------------------------------------------------------------------------------------------"
        SendKeys "{ENTER}"
        SendKeys "Select Case menusel"
        SendKeys "{ENTER}"
                         rst2.MoveFirst
                         VariableIndex = 0
                         For i = 0 To List1.ListCount - 1
                         Select Case rst2![Caption]
                         Case "/separator/" 'separator = do nothing as you can't select it
                         Case Else
                            If rst2![ItemSubMenu] = 0 Then
                                    SendKeys "{ENTER}"
                                    SendKeys "Case " & rst2![ItemMagicNum] & "'{(}" & rst2![CompletePath] & "{)}"
                                    'now what to do ?
                                    SendKeys "{ENTER}"
                                    SendKeys "MsgBox" & Chr(34) & rst2![CompletePath] & " Clicked!" & Chr(34)
                                    
                                    'if it is marked checked
                                    
                                    If rst2![Checked] = "MFS_CHECKED" Or rst2![Checked] = "MFS_UNCHECKED" Then
                                    markmessage = True
                                    VariableIndex = VariableIndex + 1
                                    SendKeys "{ENTER}"
                                    SendKeys "If APICheck" & VariableIndex & " = True Then"
                                    SendKeys "{ENTER}"
                                    SendKeys "APICheck" & VariableIndex & " = False"
                                    SendKeys "{ENTER}"
                                    SendKeys "Else"
                                    SendKeys "{ENTER}"
                                    SendKeys "APICheck" & VariableIndex & " = True"
                                    SendKeys "{ENTER}"
                                    SendKeys "End If"
                                    SendKeys "{ENTER}"
                                    End If
    
                             End If
                         End Select
                         rst2.MoveNext
                         Next i
       SendKeys "{ENTER}"
       SendKeys "Case Else"
       SendKeys "{ENTER}"
       SendKeys "End Select"
       
100
       SendKeys "{ENTER}"


DoEvents 'otherwise all wrong ! part of upper stuf goes to the wrong textbox

'message procedure
If markmessage = True Or PicMsgBoxWarning = True Then


If PicMsgBoxWarning = True Then
PicMessageForm.Show
PicMessageForm.Text1.SetFocus
SendKeys "'You have choosen to have some pics on your menu:"
SendKeys "{ENTER}"

rst2.MoveFirst
picindex = 0
For i = 0 To List1.ListCount - 1
        If rst2![Picture] = "Yes" Then
        If Mid(List2.List(i), 2, 1) = "P" Then
                SendKeys "Place pic control on form named: MenuPicContainer{(}" & picindex & "{)}"
                SendKeys "{ENTER}"
                SendKeys "Set its's picture prop. to picture that you want to appear beside item called: " & rst2![CompletePath] & " ,When Checked"
                SendKeys "{ENTER}"
                picindex = picindex + 1
                SendKeys "Place pic control on form named: MenuPicContainer{(}" & picindex & "{)}"
                SendKeys "{ENTER}"
                SendKeys "Set its's picture prop. to picture that you want to appear beside item called: " & rst2![CompletePath] & " ,When UNChecked"
                SendKeys "{ENTER}"
                picindex = picindex + 1
                Else
                SendKeys "Place pic control on form named: MenuPicContainer{(}" & picindex & "{)}"
                SendKeys "{ENTER}"
                SendKeys "Set its's picture prop. to picture that you want to appear beside item called: " & rst2![CompletePath]
                SendKeys "{ENTER}"
                picindex = picindex + 1
                End If
        End If
rst2.MoveNext
Next i
End If

'if you have some items checked
If markmessage = True Then
markmessage = False
PicMessageForm.Show
PicMessageForm.Text1.SetFocus
SendKeys "{ENTER}"
SendKeys "'Here is what you have to do manualy, because you included some checkmarks:"
SendKeys "{ENTER}"
SendKeys "{ENTER}"
SendKeys "'Under generall section of your form place:"
SendKeys "{ENTER}"


'general section of form
rst2.MoveFirst
VariableIndex = 0
For i = 0 To List1.ListCount - 1
      If rst2![Checked] = "MFS_CHECKED" Or rst2![Checked] = "MFS_UNCHECKED" Then
      VariableIndex = VariableIndex + 1
      SendKeys "Dim APICheck" & VariableIndex & " as Boolean"
      SendKeys "{ENTER}"
      End If
rst2.MoveNext
Next i

SendKeys "{ENTER}"
SendKeys "'Under Form_Load event place:"
SendKeys "{ENTER}"

'form_load event
rst2.MoveFirst
VariableIndex = 0
For i = 0 To List1.ListCount - 1
      If rst2![Checked] = "MFS_CHECKED" Then
      VariableIndex = VariableIndex + 1
      SendKeys "APICheck" & VariableIndex & " = True"
      SendKeys "{ENTER}"
      ElseIf rst2![Checked] = "MFS_UNCHECKED" Then
      VariableIndex = VariableIndex + 1
      SendKeys "APICheck" & VariableIndex & " = False"
      SendKeys "{ENTER}"
      End If
rst2.MoveNext
Next i
End If

SendKeys "{ENTER}"
SendKeys "Now that computer did big work for you, you should stand up and fly to Bahamas."
SendKeys "{ENTER}"
SendKeys "I wish you sunny weather."

End If
'end of message procedure


'and on the end, dbs closing stuff
Set rst = Nothing
Set rst1 = Nothing
Set rst2 = Nothing

dbs.Close

'next one solves hung up of this sub in case that you run it
'again while stuff is already in text2, forgot how  = did no comment, but it works
controler = 0



End Sub

Private Sub Command10_Click()
Timer1.Enabled = False
DoEvents
ApiPopUpMenuCodeGenerator
End Sub

Private Sub Command11_Click() 'exit
End
End Sub

Private Sub Command12_Click() 'help
Timer1.Enabled = False
MsgBox "This App is here to cut tons of time needed to build Api PopUp Menu. Hmm isn't Vb menu just as good ? No. First and most important VB menu can't exist on borderless form , allso Api menu has more options." & Chr(10) & Chr(10) & "If you ever build menu in VB, just click somewhere in Structure Building Box. PopUpMenu build by this wery app will appear and everything else will be obvious. For start you can click Load Structure to load TestPro.map and get into it quicker. Then:" & Chr(10) & Chr(10) & "1.Open new VB project, add module" & Chr(10) & "2.Into module paste code builded with  'Public Declares' button" & Chr(10) & "3.Under Form_Click event paste code builded by 'Generate Code' button" & Chr(10) & "4.Run new project, click anywhere in form", , "MC API Menu Code Generator ver 2.0 "
End Sub



Private Sub Command2_Click()
MsgBox "I'm Looking for sample that puts item in forms sys menu, AND THEN ACTUALY DO SOMETHING CLICKING ON IT. (kozlicki@yahoo.com) Thanks."
End Sub

Private Sub Command3_Click()
CommonDialog1.ShowSave
If CommonDialog1.FileTitle = "" Then Exit Sub 'error handler
Open CommonDialog1.FileName For Output As #1    ' Open file for output.
'lists contence
For i = 0 To List1.ListCount - 1
Write #1, List1.List(i), List2.List(i), List3.List(i), List4.List(i)
Next i

'And general settings
Write #1, "GS", "", "", ""
If Option6(0).Value = True Then Write #1, "TPM_RIGHTALIGN = &H8&", "", "", ""
If Option6(1).Value = True Then Write #1, "TPM_LEFTALIGN = &H0", "", "", ""
If Option6(2).Value = True Then Write #1, "TPM_CENTERALIGN = &H4&", "", "", ""

If Option7(0).Value = True Then Write #1, "TPM_RIGHTBUTTON = &H2&", "", "", ""
If Option7(1).Value = True Then Write #1, "TPM_LEFTBUTTON = &H0", "", "", ""

Close #1   ' Close file.

End Sub

Private Sub Command4_Click()
'erase if anything already there
List1.Clear
List2.Clear
List3.Clear
List4.Clear

On Error GoTo errorhandler
CommonDialog1.ShowOpen
Open CommonDialog1.FileName For Input As #1    ' Open file for output.
    Me.Caption = "[" & CommonDialog1.FileTitle & "]  " & "MC API Menu Code Generator ver 2.0 (Pro)"
    i = 0
    Do
    Input #1, a, b, c, d
    If a = "GS" Then Exit Do
    List1.AddItem a
    List2.AddItem b
    List3.AddItem c
    List4.AddItem d
    i = i + 1
    Loop
    
    Input #1, a, b, c, d
    
    If a = "TPM_RIGHTALIGN = &H8&" Then Option6(0).Value = True
    If a = "TPM_LEFTALIGN = &H0" Then Option6(1).Value = True
    If a = "TPM_CENTERALIGN = &H4&" Then Option6(2).Value = True
    
    Input #1, a, b, c, d

    If a = "TPM_RIGHTBUTTON = &H2&" Then Option7(0).Value = True
    If a = "TPM_LEFTBUTTON = &H0" Then Option7(1).Value = True
    
Close #1   ' Close file.
Exit Sub
errorhandler:
Select Case Err
Case 62
'appears always
Case Else
End Select
Close #1   ' Close file.

End Sub








Private Sub Command5_Click()
MsgBox "Last second discovered: THERE ARE NO BUGS !" & Chr(10) & "I PLEAD NOT GUILTY on next matters: NONLOGICAL MENU STRUCTURE, Some properties of items may exclude others i.e. U can't have Checkmark style RADIO BUTTON and same time Pic marked as Yes. Pic owerrides RADIO BUTTON, etc. If substantial voting, I might add this errors finder in future.", , "MC API Menu Code Generator ver 2.0 "
End Sub

Private Sub Command6_Click()
'run app that we will send commands to
Shell GetWindowsDirectory() & "notepad " & App.Path & "\Documentation.txt", vbMaximizedFocus
End Sub

Private Sub Command7_Click() 'close general settings
Frame7.Visible = False
End Sub

Public Sub Command8_Click()

'code to fit into module

Text2.SetFocus
Text2.Text = ""

'1.declaration section
SendKeys "'Declaration section"
SendKeys "{ENTER}"
SendKeys "Public Declare Function CreatePopupMenu Lib" & Chr(34) & "user32.dll" & Chr(34) & " {(}{)}  As Long"
SendKeys "{ENTER}"
SendKeys "Public Declare Function DestroyMenu Lib " & Chr(34) & "user32.dll" & Chr(34) & " {(}ByVal hMenu As Long{)} As Long"
SendKeys "{ENTER}"
SendKeys "Public Type MENUITEMINFO"
SendKeys "{ENTER}"
SendKeys "        cbSize As Long"
SendKeys "{ENTER}"
SendKeys "        fMask As Long"
SendKeys "{ENTER}"
SendKeys "        fType As Long"
SendKeys "{ENTER}"
SendKeys "        fState As Long"
SendKeys "{ENTER}"
SendKeys "        wID As Long"
SendKeys "{ENTER}"
SendKeys "        hSubMenu As Long"
SendKeys "{ENTER}"
SendKeys "        hbmpChecked As Long"
SendKeys "{ENTER}"
SendKeys "        hbmpUnchecked As Long"
SendKeys "{ENTER}"
SendKeys "        dwItemData As Long"
SendKeys "{ENTER}"
SendKeys "        dwTypeData As String"
SendKeys "{ENTER}"
SendKeys "        cch As Long"
SendKeys "{ENTER}"
SendKeys "End Type"
SendKeys "{ENTER}"

'Constant Definitions
 
SendKeys "Public Const MIIM_STATE = &H1"
SendKeys "{ENTER}"
SendKeys "Public Const MIIM_ID = &H2"
SendKeys "{ENTER}"
SendKeys "Public Const MIIM_SUBMENU = &H4"
SendKeys "{ENTER}"
SendKeys "Public Const MIIM_CHECKMARKS = &H8"
SendKeys "{ENTER}"
SendKeys "Public Const MIIM_DATA = &H20"
SendKeys "{ENTER}"
SendKeys "Public Const MIIM_TYPE = &H10"
SendKeys "{ENTER}"
SendKeys "Public Const MFT_BITMAP = &H4"
SendKeys "{ENTER}"
SendKeys "Public Const MFT_MENUBARBREAK = &H20"
SendKeys "{ENTER}"
SendKeys "Public Const MFT_MENUBREAK = &H40"
SendKeys "{ENTER}"
SendKeys "Public Const MFT_OWNERDRAW = &H100"
SendKeys "{ENTER}"
SendKeys "Public Const MFT_RADIOCHECK = &H200"
SendKeys "{ENTER}"
SendKeys "Public Const MFT_RIGHTJUSTIFY = &H4000"
SendKeys "{ENTER}"
SendKeys "Public Const MFT_RIGHTORDER = &H2000"
SendKeys "{ENTER}"
SendKeys "Public Const MFT_SEPARATOR = &H800"
SendKeys "{ENTER}"
SendKeys "Public Const MFT_STRING = &H0"
SendKeys "{ENTER}"
SendKeys "Public Const MFS_CHECKED = &H8"
SendKeys "{ENTER}"
SendKeys "Public Const MFS_DEFAULT = &H1000"
SendKeys "{ENTER}"
SendKeys "Public Const MFS_DISABLED = &H2"
SendKeys "{ENTER}"
SendKeys "Public Const MFS_ENABLED = &H0"
SendKeys "{ENTER}"
SendKeys "Public Const MFS_GRAYED = &H1"
SendKeys "{ENTER}"
SendKeys "Public Const MFS_HILITE = &H80"
SendKeys "{ENTER}"
SendKeys "Public Const MFS_UNCHECKED = &H0"
SendKeys "{ENTER}"
SendKeys "Public Const MFS_UNHILITE = &H0"

'functions = API-s
SendKeys "{ENTER}"
SendKeys "Public Declare Function InsertMenuItem Lib " & Chr(34) & "user32.dll" & Chr(34) & " Alias " & Chr(34) & "InsertMenuItemA" & Chr(34) & " _"
SendKeys "{ENTER}"
SendKeys "{(}ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO{)} As Long"
SendKeys "{ENTER}"
SendKeys "Public Declare Function TrackPopupMenu Lib " & Chr(34) & "user32.dll" & Chr(34) & " _"
SendKeys "{ENTER}"
SendKeys "{(}ByVal hMenu As Long, ByVal uFlags As Long, ByVal x As Long, ByVal y As Long," & " _"
SendKeys "{ENTER}"
SendKeys "ByVal nReserved As Long, ByVal hWnd As Long, ByVal prcRect As Long{)} As Long"
SendKeys "{ENTER}"
SendKeys "Public Const TPM_RIGHTALIGN = &H8&"
SendKeys "{ENTER}"
SendKeys "Public Const TPM_CENTERALIGN = &H4&"
SendKeys "{ENTER}"
SendKeys "Public Const TPM_LEFTALIGN = &H0"
SendKeys "{ENTER}"
SendKeys "Public Const TPM_TOPALIGN = &H0"
SendKeys "{ENTER}"
SendKeys "Public Const TPM_NONOTIFY = &H80"
SendKeys "{ENTER}"
SendKeys "Public Const TPM_RETURNCMD = &H100"
SendKeys "{ENTER}"
SendKeys "Public Const TPM_LEFTBUTTON = &H0"
SendKeys "{ENTER}"
SendKeys "Public Const  TPM_RIGHTBUTTON = &H2&"
SendKeys "{ENTER}"
SendKeys "Public Type POINT_TYPE"
SendKeys "{ENTER}"
SendKeys "x As Long"
SendKeys "{ENTER}"
SendKeys "y As Long"
SendKeys "{ENTER}"
SendKeys "End Type"
SendKeys "{ENTER}"
SendKeys "Public Declare Function GetCursorPos Lib " & Chr(34) & "user32.dll" & Chr(34) & " {(}lpPoint As POINT_TYPE{)} As Long"
SendKeys "{ENTER}"
SendKeys "Public Declare Function AppendMenu Lib " & Chr(34) & "user32" & Chr(34) & " Alias " & Chr(34) & "AppendMenuA" & Chr(34) & " {(}ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any{)} As Long"
SendKeys "{ENTER}" 'for adding pictures
SendKeys "Public Declare Function SetMenuItemBitmaps Lib " & Chr(34) & "user32" & Chr(34) & " {(}ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long{)} As Long"




End Sub

Private Sub Command9_Click()

Clipboard.Clear
Clipboard.SetText Text2.Text, vbCFText
End Sub





Private Sub Form_Activate()
'MsgBox " To get all work, do please extract all files from zip. Especialy My.mdb is absolutely necesary.Otherwise nothing works.", , "MC API Menu Code Generator ver 2.0 "

End Sub


Private Sub Label4_Click()
Open App.Path & "\InetAdress.txt" For Input As #1    ' Open file for output.
Input #1, a
Close #1
Shell "start " & a
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ClickNoCanDo = True Then Exit Sub
'teh following lines trigger some click events so inhibit that...
ClickNoCanDo = True
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
ClickNoCanDo = False


Dim hPopupMenu1 As Long ' handle to the popup menu to display
Dim hPopupMenu2 As Long ' handle to the popup menu to display
Dim hPopupMenu3 As Long ' handle to the popup menu to display
Dim mii1 As MENUITEMINFO   ' describes menu items to add
Dim mii2 As MENUITEMINFO   ' describes menu items to add
Dim mii3 As MENUITEMINFO   ' describes menu items to add
Dim curpos As POINT_TYPE  ' holds the current mouse coordinates
Dim menusel As Long       ' ID of what the user selected in the popup menu
Dim retval As Long        ' generic return value


'Create the popup menus which are initialy empty.
hPopupMenu1 = CreatePopupMenu()
hPopupMenu2 = CreatePopupMenu()
hPopupMenu3 = CreatePopupMenu()

'Create the structure which is the base for all menus:
With mii1
.cbSize = Len(mii1) ' The size of this structure.
.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
End With

'Make all structures equal
mii2 = mii1
mii3 = mii1

With mii1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1001 ' Assign this item an item identifier.
.dwTypeData = "Item"
.cch = Len("Item")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 0, 1, mii1)

With mii1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1000 ' Assign this item an item identifier.
.dwTypeData = "Separator"
.cch = Len("Separator")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 1, 1, mii1)

With mii2
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1003 ' Assign this item an item identifier.
.dwTypeData = "Item"
.cch = Len("Item")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 0, 1, mii2)

With mii2
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1002 ' Assign this item an item identifier.
.dwTypeData = "Separator"
.cch = Len("Separator")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 1, 1, mii2)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1015 ' Assign this item an item identifier.
.dwTypeData = "Move Right"
.cch = Len("Move Right")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 0, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1014 ' Assign this item an item identifier.
.dwTypeData = "Move Left"
.cch = Len("Move Left")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 1, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1013 ' Assign this item an item identifier.
.dwTypeData = "Move Up"
.cch = Len("Move Up")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 2, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1012 ' Assign this item an item identifier.
.dwTypeData = "Move Down"
.cch = Len("Move Down")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 3, 1, mii3)

With mii3
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1011 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 4, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1010 ' Assign this item an item identifier.
.dwTypeData = "Menu Behaviour"
.cch = Len("Menu Behaviour")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 5, 1, mii3)

With mii3
.fType = MFT_STRING Or MFT_MENUBARBREAK
.fState = MFS_ENABLED
.wID = 1009 ' Assign this item an item identifier.
.dwTypeData = "Add"
.cch = Len("Add")
.hSubMenu = hPopupMenu2
End With
retval = InsertMenuItem(hPopupMenu3, 6, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1008 ' Assign this item an item identifier.
.dwTypeData = "Insert"
.cch = Len("Insert")
.hSubMenu = hPopupMenu1
End With
retval = InsertMenuItem(hPopupMenu3, 7, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1007 ' Assign this item an item identifier.
.dwTypeData = "Delete"
.cch = Len("Delete")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 8, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1006 ' Assign this item an item identifier.
.dwTypeData = "Change String"
.cch = Len("Change String")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 9, 1, mii3)

With mii3
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1005 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 10, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1004 ' Assign this item an item identifier.
.dwTypeData = "Close Menu"
.cch = Len("Item properties")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 11, 1, mii3)

'The following code is for adding pictures into menus!
'------------------------------------------------------------
'------------------------------------------------------------

retval = SetMenuItemBitmaps(hPopupMenu3, 1015, 1, MenuPicContainer(0), MenuPicContainer(0))
retval = SetMenuItemBitmaps(hPopupMenu3, 1014, 1, MenuPicContainer(1), MenuPicContainer(1))
retval = SetMenuItemBitmaps(hPopupMenu3, 1013, 1, MenuPicContainer(2), MenuPicContainer(2))
retval = SetMenuItemBitmaps(hPopupMenu3, 1012, 1, MenuPicContainer(3), MenuPicContainer(3))
'------------------------------------------------------------
'------------------------------------------------------------

retval = GetCursorPos(curpos)
menusel = TrackPopupMenu(hPopupMenu3, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_CENTERALIGN Or TPM_RIGHTBUTTON, curpos.X, curpos.Y, 0, Form1.hWnd, 0)
retval = DestroyMenu(hPopupMenu3)
'------------------------------------------------------------------------------------------------
'DOWN BELOW  PUT IN YOUR CODE MANUALY !!!!
'------------------------------------------------------------------------------------------------
Select Case menusel

Case 1001 '(Insert/Item)

'user errors .......
                If List1.ListCount = 0 Then
                MsgBox "Can't insert as there is nothing in list !"
                Exit Sub
                End If
                
                Dim default
                default = ""
                message = "Type in new item caption"
                Title = "Input box"
                myvalue = InputBox(message, Title, default)
                If myvalue <> "" Then
                List1.AddItem myvalue, List1.ListIndex
                List2.AddItem "S  ", List1.ListIndex 'MFT_STRING
                List3.AddItem "EN ", List1.ListIndex 'MFS_ENABLED
                List4.AddItem "No", List1.ListIndex
                End If

Case 1000 '(Insert/Separator)

'user errors .......
                If List1.ListCount = 0 Then
                MsgBox "Can't insert as there is nothing in list !"
                Exit Sub
                End If
List1.AddItem "/separator/", List1.ListIndex
List2.AddItem "S  ", List1.ListIndex
List3.AddItem "EN ", List1.ListIndex
List4.AddItem "No", List1.ListIndex

Case 1003 '(Add/Item)

 'Dim default
            default = ""
            message = "Type in new item caption"
            Title = "Input box"
            myvalue = InputBox(message, Title, default)
            
            If myvalue <> "" Then
            List1.AddItem myvalue
            List2.AddItem "S  "
            List3.AddItem "EN "
            List4.AddItem "No"
            End If

Case 1002 '(Add/Separator)

'user errors .......
                If List1.ListCount = 0 Then
                MsgBox "Having separator on top of menu is a dumb thing. Can't permit that !"
                Exit Sub
                End If
List1.AddItem "/separator/"
List2.AddItem "S  "
List3.AddItem "EN "
List4.AddItem "No"

Case 1015 '(Move Right)
If List1.ListCount = 0 Then Exit Sub
List1.List(List1.ListIndex) = "----" & List1.List(List1.ListIndex)
Case 1014 '(Move Left)
If List1.ListCount = 0 Then Exit Sub
If Left(List1.List(List1.ListIndex), 1) = "-" Then
List1.List(List1.ListIndex) = Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 4)
End If
Case 1013 '(Move Up)
If List1.ListCount = 0 Then Exit Sub
If List1.ListIndex > 0 Then
List1.Selected(List1.ListIndex - 1) = True 'highlight one item up
End If
Case 1012 '(Move Down)
If List1.ListCount = 0 Then Exit Sub
If List1.ListIndex < List1.ListCount - 1 Then
List1.Selected(List1.ListIndex + 1) = True 'highlight one item down
End If

Case 1010 '(Menu Behaviour)

Frame7.Visible = True

Case 1007 '(Delete)

 'user errors .......
                If List1.ListCount = 0 Then
                MsgBox "You can't delete as there is nothing to delete !"
                Exit Sub
                End If
                If List1.List(List1.ListIndex) = "" Then
                MsgBox "No item selected !"
                Exit Sub
                End If
        List4.RemoveItem List1.ListIndex 'Action
        List2.RemoveItem List1.ListIndex 'Action
        List3.RemoveItem List1.ListIndex 'Action
        List1.RemoveItem List1.ListIndex 'Action

Case 1006 '(ChangeString)

If List1.ListIndex = -1 Then MsgBox "Can't change as there no selection": Exit Sub 'on empty list, or nozhing selected
StrA = InputBox("Enter new item string to replace current one", , List1.List(List1.ListIndex))
If StrA = "" Then Exit Sub
List1.List(List1.ListIndex) = StrA

Case 1004 '(Close Menu)
'nothing as just closing menu

Case Else 'from select case menusel , far up
End Select

End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List2.ToolTipText = List2.List(List2.ListIndex)
End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List3.ToolTipText = List3.List(List3.ListIndex)
End Sub

Private Sub List4_Click()
If ClickNoCanDo = True Then Exit Sub
'teh following lines triggere some click events so inhibit that...
ClickNoCanDo = True
List1.ListIndex = List4.ListIndex
List2.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
ClickNoCanDo = False

Dim hPopupMenu1 As Long ' handle to the popup menu to display
Dim mii1 As MENUITEMINFO   ' describes menu items to add
Dim curpos As POINT_TYPE  ' holds the current mouse coordinates
Dim menusel As Long       ' ID of what the user selected in the popup menu
Dim retval As Long        ' generic return value


'Create the popup menus which are initialy empty.
hPopupMenu1 = CreatePopupMenu()

'Create the structure which is the base for all menus:
With mii1
.cbSize = Len(mii1) ' The size of this structure.
.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
End With

'Make all structures equal

With mii1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1004 ' Assign this item an item identifier.
.dwTypeData = "Yes"
.cch = Len("Yes")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 0, 1, mii1)

With mii1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1003 ' Assign this item an item identifier.
.dwTypeData = "No"
.cch = Len("No")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 1, 1, mii1)

With mii1
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1002 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 2, 1, mii1)

With mii1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1001 ' Assign this item an item identifier.
.dwTypeData = "Yes To All"
.cch = Len("Yes To All")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 3, 1, mii1)

With mii1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1000 ' Assign this item an item identifier.
.dwTypeData = "No To All"
.cch = Len("No To All")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 4, 1, mii1)

'The following code is for adding pictures into menus!
'------------------------------------------------------------
'------------------------------------------------------------

'------------------------------------------------------------
'------------------------------------------------------------

retval = GetCursorPos(curpos)
menusel = TrackPopupMenu(hPopupMenu1, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_RIGHTALIGN Or TPM_RIGHTBUTTON, curpos.X, curpos.Y, 0, Form1.hWnd, 0)
retval = DestroyMenu(hPopupMenu1)
'------------------------------------------------------------------------------------------------
'DOWN BELOW  PUT IN YOUR CODE MANUALY !!!!
'------------------------------------------------------------------------------------------------
Select Case menusel

Case 1004 '(Yes)
        If List1.List(List4.ListIndex) = "/separator/" Then MsgBox " Separators can't have pictures !": Exit Sub
        List4.List(List4.ListIndex) = "Yes"
Case 1003 '(No)
List4.List(List4.ListIndex) = "No"

Case 1001 '(Yes To All)
        For i = 0 To List4.ListCount - 1
            If List1.List(i) <> "/separator/" Then
            List4.List(i) = "Yes"
            End If
        Next i
Case 1000 '(No To All)
        For i = 0 To List4.ListCount - 1
        List4.List(i) = "No"
        Next i
Case Else
End Select
End Sub




Private Sub Option1_Click(Index As Integer)
If Index = 1 Or Index = 0 Then
Check1.Enabled = True
Else
Check1.Enabled = False
End If
End Sub

Private Sub Form_Load()
Text2.Width = Screen.Width
End Sub


Private Sub List2_Click()
If ClickNoCanDo = True Then Exit Sub
'teh following lines triggere some click events so inhibit that...
ClickNoCanDo = True
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List4.ListIndex = List2.ListIndex
ClickNoCanDo = False

'and menu to be opened stuff
'--------------------------------------------------------------------------------------------------------------------------
'CODE AUTOGENERATED WITH:  MC API Menu Code Generator ver 2.0
'---------------------------------------------------------------------------------------------------------------------------
Dim hPopupMenu1 As Long ' handle to the popup menu to display
Dim hPopupMenu2 As Long ' handle to the popup menu to display
Dim hPopupMenu3 As Long ' handle to the popup menu to display
Dim mii1 As MENUITEMINFO   ' describes menu items to add
Dim mii2 As MENUITEMINFO   ' describes menu items to add
Dim mii3 As MENUITEMINFO   ' describes menu items to add
Dim curpos As POINT_TYPE  ' holds the current mouse coordinates
Dim menusel As Long       ' ID of what the user selected in the popup menu
Dim retval As Long        ' generic return value


'Create the popup menus which are initialy empty.
hPopupMenu1 = CreatePopupMenu()
hPopupMenu2 = CreatePopupMenu()
hPopupMenu3 = CreatePopupMenu()

'Create the structure which is the base for all menus:
With mii1
.cbSize = Len(mii1) ' The size of this structure.
.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
End With

'Make all structures equal
mii2 = mii1
mii3 = mii1

With mii1 '(Insert column break/With dividing line)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1002 ' Assign this item an item identifier.
.dwTypeData = "With dividing line"
.cch = Len("With dividing line")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 0, 1, mii1)

With mii1 '(Insert column break/Without dividing line)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1001 ' Assign this item an item identifier.
.dwTypeData = "Without dividing line"
.cch = Len("Without dividing line")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 1, 1, mii1)

With mii1 '(Insert column break/Delete col. break)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1000 ' Assign this item an item identifier.
.dwTypeData = "Delete col. break"
.cch = Len("Delete col. break")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 2, 1, mii1)

With mii2 '(Check style/None)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1006 ' Assign this item an item identifier.
.dwTypeData = "None"
.cch = Len("None")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 0, 1, mii2)

With mii2 '(Check style/Normal checkmark)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1005 ' Assign this item an item identifier.
.dwTypeData = "Normal checkmark"
.cch = Len("Normal checkmark")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 1, 1, mii2)

With mii2 '(Check style/Radio button)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1004 ' Assign this item an item identifier.
.dwTypeData = "Radio button"
.cch = Len("Radio button")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 2, 1, mii2)

With mii2 '(Check style/Pictures)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1003 ' Assign this item an item identifier.
.dwTypeData = "Pictures"
.cch = Len("Pictures")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 3, 1, mii2)

With mii3 '(Check style)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1010 ' Assign this item an item identifier.
.dwTypeData = "Check style"
.cch = Len("Check style")
.hSubMenu = hPopupMenu2
End With
retval = InsertMenuItem(hPopupMenu3, 0, 1, mii3)

With mii3 '(Insert column break)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1009 ' Assign this item an item identifier.
.dwTypeData = "Insert column break"
.cch = Len("Insert column break")
.hSubMenu = hPopupMenu1
End With
retval = InsertMenuItem(hPopupMenu3, 1, 1, mii3)

With mii3 '(Help !)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1008 ' Assign this item an item identifier.
.dwTypeData = "Help !"
.cch = Len("Help !")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 2, 1, mii3)

With mii3 '(Close Menu)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1007 ' Assign this item an item identifier.
.dwTypeData = "Close Menu"
.cch = Len("Close Menu")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 3, 1, mii3)

'The following code is for adding pictures into menus, if there are any!
'------------------------------------------------------------
'------------------------------------------------------------

'------------------------------------------------------------
'------------------------------------------------------------

retval = GetCursorPos(curpos)
menusel = TrackPopupMenu(hPopupMenu3, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_RIGHTALIGN Or TPM_RIGHTBUTTON, curpos.X, curpos.Y, 0, Form1.hWnd, 0)
retval = DestroyMenu(hPopupMenu3)
'------------------------------------------------------------------------------------------------
'DOWN BELOW  PUT IN YOUR CODE MANUALY !!!!
'------------------------------------------------------------------------------------------------
Dim IncomingString As String
Dim outputstring As String
IncomingString = List2.List(List2.ListIndex)

Select Case menusel

Case 1002 '(Insert column break/With dividing line)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "L"
        List2.List(List2.ListIndex) = outputstring
Case 1001 '(Insert column break/Without dividing line)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "N"
        List2.List(List2.ListIndex) = outputstring
Case 1000 '(Insert column break/Delete col. break)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & " "
        List2.List(List2.ListIndex) = outputstring
Case 1006 '(Check style/None)
        outputstring = Mid(IncomingString, 1, 1) & " " & Mid(IncomingString, 3, 1)
        List2.List(List2.ListIndex) = outputstring
Case 1005 '(Check style/Normal checkmark)
        outputstring = Mid(IncomingString, 1, 1) & "C" & Mid(IncomingString, 3, 1)
        List2.List(List2.ListIndex) = outputstring
        If Mid(List3.List(List2.ListIndex), 3, 1) = " " Then
        MsgBox "Go to Item state list box (at this item) and tell me what will be the start state of this item, there select Check Opt ?/Yes and then by your mind "
        End If
Case 1004 '(Check style/Radio button)
        outputstring = Mid(IncomingString, 1, 1) & "R" & Mid(IncomingString, 3, 1)
        List2.List(List2.ListIndex) = outputstring
        If Mid(List3.List(List2.ListIndex), 3, 1) = " " Then
        MsgBox "Go to Item state list box (at this item) and tell me what will be the start state of this item, there select Check Opt ?/Yes and then by your mind "
        End If
Case 1003 '(Check style/Pictures)
        outputstring = Mid(IncomingString, 1, 1) & "P" & Mid(IncomingString, 3, 1)
        'there comes a package
        List2.List(List2.ListIndex) = outputstring
        List4.List(List2.ListIndex) = "Yes"
        If Mid(List3.List(List2.ListIndex), 3, 1) = " " Then
        MsgBox "Go to Item state list box (at this item) and tell me what will be the start state of this item, there select Check Opt ?/Yes and then by your mind "
        End If
Case 1008 '(Help !)
        MsgBox "1) S = String(MFT_STRING), that is default, must be here in any case" & Chr(10) & _
        "2) R = RadioButton(MFT_RADIOCHECK), type of checkmark" & Chr(10) & _
        "    C = Normal checkmark(nothinga as it is default type of checkmark)" & Chr(10) & _
        "    P = Pictures(nothing - you will use 2 different .bmp)" & Chr(10) & _
        "3) L = Column break with dividing Line(MFT_MENUBARBREAK)" & Chr(10) & _
        "    N = Column break without dividing Line(MFT_MENUBREAK)", , "MC API Menu Code Generator ver 2.0 "

Case 1007 '(Close Menu)
'do nothing
Case Else
End Select






End Sub

Private Sub List3_Click()
If ClickNoCanDo = True Then Exit Sub
'teh following lines triggere some click events so inhibit that...
ClickNoCanDo = True
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List4.ListIndex = List3.ListIndex
ClickNoCanDo = False
'and menu to be opened stuff
'--------------------------------------------------------------------------------------------------------------------------
'CODE AUTOGENERATED WITH:  MC API Menu Code Generator ver 2.0
'---------------------------------------------------------------------------------------------------------------------------
Dim hPopupMenu1 As Long ' handle to the popup menu to display
Dim hPopupMenu2 As Long ' handle to the popup menu to display
Dim hPopupMenu3 As Long ' handle to the popup menu to display
Dim hPopupMenu4 As Long ' handle to the popup menu to display
Dim hPopupMenu5 As Long ' handle to the popup menu to display
Dim mii1 As MENUITEMINFO   ' describes menu items to add
Dim mii2 As MENUITEMINFO   ' describes menu items to add
Dim mii3 As MENUITEMINFO   ' describes menu items to add
Dim mii4 As MENUITEMINFO   ' describes menu items to add
Dim mii5 As MENUITEMINFO   ' describes menu items to add
Dim curpos As POINT_TYPE  ' holds the current mouse coordinates
Dim menusel As Long       ' ID of what the user selected in the popup menu
Dim retval As Long        ' generic return value


'Create the popup menus which are initialy empty.
hPopupMenu1 = CreatePopupMenu()
hPopupMenu2 = CreatePopupMenu()
hPopupMenu3 = CreatePopupMenu()
hPopupMenu4 = CreatePopupMenu()
hPopupMenu5 = CreatePopupMenu()

'Create the structure which is the base for all menus:
With mii1
.cbSize = Len(mii1) ' The size of this structure.
.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
End With

'Make all structures equal
mii2 = mii1
mii3 = mii1
mii4 = mii1
mii5 = mii1

With mii1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1001 ' Assign this item an item identifier.
.dwTypeData = "On Start Checked"
.cch = Len("On Start Checked")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 0, 1, mii1)

With mii1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1000 ' Assign this item an item identifier.
.dwTypeData = "On Start Unchecked"
.cch = Len("On Start Unchecked")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 1, 1, mii1)

With mii2
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1003 ' Assign this item an item identifier.
.dwTypeData = "On Start Checked"
.cch = Len("On Start Checked")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 0, 1, mii2)

With mii2
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1002 ' Assign this item an item identifier.
.dwTypeData = "On Start Unchecked"
.cch = Len("On Start Unchecked")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 1, 1, mii2)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1005 ' Assign this item an item identifier.
.dwTypeData = "No1"
.cch = Len("No1")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 0, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1004 ' Assign this item an item identifier.
.dwTypeData = "Yes1"
.cch = Len("Yes1")
.hSubMenu = hPopupMenu1
End With
retval = InsertMenuItem(hPopupMenu3, 1, 1, mii3)

With mii4
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1007 ' Assign this item an item identifier.
.dwTypeData = "No"
.cch = Len("No")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu4, 0, 1, mii4)

With mii4
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1006 ' Assign this item an item identifier.
.dwTypeData = "Yes"
.cch = Len("Yes")
.hSubMenu = hPopupMenu2
End With
retval = InsertMenuItem(hPopupMenu4, 1, 1, mii4)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT Or MFS_GRAYED
.wID = 1029 ' Assign this item an item identifier.
.dwTypeData = "Bold All"
.cch = Len("Bold All")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 0, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_GRAYED
.wID = 1028 ' Assign this item an item identifier.
.dwTypeData = "Normal All"
.cch = Len("Normal All")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 1, 1, mii5)

With mii5
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1027 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 2, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_GRAYED
.wID = 1026 ' Assign this item an item identifier.
.dwTypeData = "Enable All"
.cch = Len("Enable All")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 3, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_GRAYED
.wID = 1025 ' Assign this item an item identifier.
.dwTypeData = "Disable All"
.cch = Len("Disable All")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 4, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_GRAYED
.wID = 1024 ' Assign this item an item identifier.
.dwTypeData = "Graye All"
.cch = Len("Gray All")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 5, 1, mii5)

With mii5
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1023 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 6, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_GRAYED
.wID = 1022 ' Assign this item an item identifier.
.dwTypeData = "Check opt All ?"
.cch = Len("Check opt All ?")
.hSubMenu = hPopupMenu4
End With
retval = InsertMenuItem(hPopupMenu5, 7, 1, mii5)

With mii5
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1021 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 8, 1, mii5)

With mii5
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1020 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 9, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1019 ' Assign this item an item identifier.
.dwTypeData = "Help !"
.cch = Len("Help !")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 10, 1, mii5)

With mii5
.fType = MFT_STRING Or MFT_MENUBARBREAK
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1018 ' Assign this item an item identifier.
.dwTypeData = "Bold"
.cch = Len("Bold")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 11, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1017 ' Assign this item an item identifier.
.dwTypeData = "Normal"
.cch = Len("Normal")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 12, 1, mii5)

With mii5
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1016 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 13, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1015 ' Assign this item an item identifier.
.dwTypeData = "Enabled"
.cch = Len("Enabled")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 14, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1014 ' Assign this item an item identifier.
.dwTypeData = "Disabled"
.cch = Len("Disabled")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 15, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1013 ' Assign this item an item identifier.
.dwTypeData = "Grayed"
.cch = Len("Grayed")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 16, 1, mii5)

With mii5
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1012 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 17, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1011 ' Assign this item an item identifier.
.dwTypeData = "Check opt ?"
.cch = Len("Check opt ?")
.hSubMenu = hPopupMenu3
End With
retval = InsertMenuItem(hPopupMenu5, 18, 1, mii5)

With mii5
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1010 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 19, 1, mii5)

With mii5
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1009 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 20, 1, mii5)

With mii5
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1008 ' Assign this item an item identifier.
.dwTypeData = "Close menu"
.cch = Len("Close menu")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 21, 1, mii5)

'The following code is for adding pictures into menus!
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------

retval = GetCursorPos(curpos)
menusel = TrackPopupMenu(hPopupMenu5, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_RIGHTALIGN Or TPM_RIGHTBUTTON, curpos.X, curpos.Y, 0, Form1.hWnd, 0)
retval = DestroyMenu(hPopupMenu5)
'------------------------------------------------------------------------------------------------
'DOWN BELOW  PUT IN YOUR CODE MANUALY !!!!
'------------------------------------------------------------------------------------------------
Dim IncomingString As String
Dim outputstring As String
IncomingString = List3.List(List3.ListIndex)


Select Case menusel

Case 1001 '(Check opt ?/Yes1/On Start Checked)

        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "C"
        List3.List(List3.ListIndex) = outputstring
        
        If Mid(List2.List(List3.ListIndex), 3, 1) = " " Then
        MsgBox "You must now select checkmark style in Itemtype listbox for this item "
        End If
Case 1000 '(Check opt ?/Yes1/On Start Unchecked)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "U"
        List3.List(List3.ListIndex) = outputstring
         
         If Mid(List2.List(List3.ListIndex), 3, 1) = " " Then
        MsgBox "You must now select checkmark style in Itemtype listbox for this item "
        End If
Case 1003 '(Check opt All ?/Yes/On Start Checked)
MsgBox "Check opt All ?/Yes/On Start Checked Clicked!"
Case 1002 '(Check opt All ?/Yes/On Start Unchecked)
MsgBox "Check opt All ?/Yes/On Start Unchecked Clicked!"
Case 1005 '(Check opt ?/No1)

        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & " "
        List3.List(List3.ListIndex) = outputstring

Case 1007 '(Check opt All ?/No)
MsgBox "Check opt All ?/No Clicked!"
Case 1029 '(Bold All)
MsgBox "Bold All Clicked!"
Case 1028 '(Normal All)
MsgBox "Normal All Clicked!"
Case 1026 '(Enable All)
MsgBox "Enable All Clicked"
Case 1025 '(Disable All)
MsgBox "Disable All Clicked!"
Case 1024 '(Gray All)
MsgBox "Graye All Clicked!"
Case 1019 '(Help !)

MsgBox "E = Enabled(MFS_ENABLED) - this is added here as default" & Chr(10) & _
"D = Disabled(MFS_DISABLED)" & Chr(10) & _
"G = Grayed(MFS_GRAYED)" & Chr(10) & _
"One of upper three on first place, default = Enabled" & Chr(10) & _
"B = Bold(MFS_DEFAULT)" & Chr(10) & _
"N = not bold(nothing as this is default)" & Chr(10) & _
"Upper two are optional on second place" & Chr(10) & _
"C = Checked(MFS_CHECKED)" & Chr(10) & _
"U = Unchecked(MFS_UNCHECHED)" & Chr(10) & _
"Upper two are optional on third place", , "MC API Menu Code Generator ver 2.0 "

Case 1018 '(Bold)
        outputstring = Mid(IncomingString, 1, 1) & "B" & Mid(IncomingString, 3, 1)
        List3.List(List3.ListIndex) = outputstring
Case 1017 '(Normal)
        outputstring = Mid(IncomingString, 1, 1) & "N" & Mid(IncomingString, 3, 1)
        List3.List(List3.ListIndex) = outputstring
Case 1015 '(Enabled)
        outputstring = "E" & Mid(IncomingString, 2, 1) & Mid(IncomingString, 3, 1)
        List3.List(List3.ListIndex) = outputstring
Case 1014 '(Disabled)
        outputstring = "D" & Mid(IncomingString, 2, 1) & Mid(IncomingString, 3, 1)
        List3.List(List3.ListIndex) = outputstring
Case 1013 '(Grayed)
        outputstring = "G" & Mid(IncomingString, 2, 1) & Mid(IncomingString, 3, 1)
        List3.List(List3.ListIndex) = outputstring
Case 1008 '(Close menu)
'MsgBox "Close menu Clicked!"
Case Else
End Select

End Sub



Private Sub Timer1_Timer()
If timervar = True Then
Command12.FontBold = True
 timervar = False
Else
Command12.FontBold = False
timervar = True
End If
End Sub
