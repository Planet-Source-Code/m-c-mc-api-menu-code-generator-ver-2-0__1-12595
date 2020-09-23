Attribute VB_Name = "MenuApis"

Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Type MENUITEMINFO
        cbSize As Long
        fMask As Long
        fType As Long
        fState As Long
        wID As Long
        hSubMenu As Long
        hbmpChecked As Long
        hbmpUnchecked As Long
        dwItemData As Long
        dwTypeData As String
        cch As Long
End Type
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_DATA = &H20
Public Const MIIM_TYPE = &H10
Public Const MFT_BITMAP = &H4
Public Const MFT_MENUBARBREAK = &H20
Public Const MFT_MENUBREAK = &H40
Public Const MFT_OWNERDRAW = &H100
Public Const MFT_RADIOCHECK = &H200
Public Const MFT_RIGHTJUSTIFY = &H4000
Public Const MFT_RIGHTORDER = &H2000
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_CHECKED = &H8
Public Const MFS_DEFAULT = &H1000
Public Const MFS_DISABLED = &H2
Public Const MFS_ENABLED = &H0
Public Const MFS_GRAYED = &H1
Public Const MFS_HILITE = &H80
Public Const MFS_UNCHECKED = &H0
Public Const MFS_UNHILITE = &H0
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" _
(ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Declare Function TrackPopupMenu Lib "user32.dll" _
(ByVal hMenu As Long, ByVal uFlags As Long, ByVal x As Long, ByVal y As Long, _
ByVal nReserved As Long, ByVal hwnd As Long, ByVal prcRect As Long) As Long
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const TPM_TOPALIGN = &H0
Public Const TPM_NONOTIFY = &H80
Public Const TPM_RETURNCMD = &H100

Public Type POINT_TYPE
x As Long
y As Long
End Type
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long


'*********************
'* Consts            *
'*********************
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Const MF_BYPOSITION = &H400&
