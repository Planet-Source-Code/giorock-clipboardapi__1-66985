VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ClipBoard API"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "&X"
      Height          =   210
      Left            =   4005
      TabIndex        =   8
      ToolTipText     =   "Clear Picture"
      Top             =   3210
      Width           =   195
   End
   Begin VB.CommandButton Command7 
      Caption         =   "CB Picture &Copy"
      Height          =   495
      Left            =   330
      TabIndex        =   7
      Top             =   5070
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CB Picture &Paste"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   5085
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CB C&lear"
      Height          =   495
      Left            =   3075
      TabIndex        =   5
      Top             =   5085
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&X"
      Height          =   210
      Left            =   4395
      TabIndex        =   4
      ToolTipText     =   "Clear List"
      Top             =   2265
      Width           =   195
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   45
      TabIndex        =   3
      Top             =   60
      Width           =   4530
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CB C&lear"
      Height          =   495
      Left            =   3075
      TabIndex        =   2
      Top             =   2430
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CB Text &Paste"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2430
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CB Text &Copy"
      Height          =   495
      Left            =   330
      TabIndex        =   0
      Top             =   2415
      Width           =   1215
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1500
      Left            =   2490
      Top             =   3210
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   705
      Picture         =   "frmClipBoardAPI.frx":0000
      Top             =   3210
      Width           =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eCBERRORMSG
    CB_OPEN_ERROR = 0
    CB_NO_BITMAP_FORMAT_AVAILABLE = 1
    CB_NO_TEXT_FORMAT_AVAILABLE = 2
    CB_ALREADY_OPEN = 3
End Enum

Private sME(3) As String

Private Declare Function GetOpenClipboardWindow Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function CountClipboardFormats Lib "user32" () As Long
Private Declare Function GetClipboardOwner Lib "user32" () As Long
Private Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ChangeClipboardChain Lib "user32" (ByVal hwnd As Long, ByVal hWndNext As Long) As Long
Private Declare Function GetClipboardViewer Lib "user32" () As Long
Private Declare Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_MOVEABLE = &H2

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2

Private Type OLEPIC
    Size As Long
    tType As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "OlePro32.dll" (PicDesc As OLEPIC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As Any) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const IMAGE_BITMAP = 0
Private Const LR_COPYRETURNORG = &H4
Private Const LR_CREATEDIBSECTION = &H2000


'Private Enum ClipBoardFormat
'    CF_ACCEPT = &H0
'    CF_ANSIONLY = &H400&
'    CF_APPLY = &H200&
'    CF_BITMAP = 2
'    CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
'    CF_BOTTOMUP_DIB = CF_DIB
'    CF_CONVERTONLY = &H100&
'    CF_DEFER = &H2
'    CF_DIB = 8
'    CF_DIBV5 = 17
'    CF_DIF = 5
'    CF_DISABLEACTIVATEAS = &H40&
'    CF_DISABLEDISPLAYASICON = &H20&
'    CF_DSPBITMAP = &H82
'    CF_DSPENHMETAFILE = &H8E
'    CF_DSPMETAFILEPICT = &H83
'    CF_DSPTEXT = &H81
'    CF_EFFECTS = &H100&
'    CF_ENABLEHOOK = &H8&
'    CF_ENABLETEMPLATE = &H10&
'    CF_ENABLETEMPLATEHANDLE = &H20&
'    CF_ENHMETAFILE = 14
'    CF_FIXEDPITCHONLY = &H4000&
'    CF_FORCEFONTEXIST = &H10000
'    CF_GDIOBJFIRST = &H300
'    CF_GDIOBJLAST = &H3FF
'    CF_HDROP = 15
'    CF_HIDECHANGEICON = &H80&
'    CF_INITTOLOGFONTSTRUCT = &H40&
'    CF_JPEG = 19
'    CF_LIMITSIZE = &H2000&
'    CF_LOCALE = 16
'    CF_MAX = 17
'    CF_METAFILEPICT = 3
'    CF_MULTI_TIFF = 22
'    CF_NOFACESEL = &H80000
'    CF_NOOEMFONTS = CF_NOVECTORFONTS
'    CF_NOSCRIPTSEL = &H800000
'    CF_NOSIMULATIONS = &H1000&
'    CF_NOSIZESEL = &H200000
'    CF_NOSTYLESEL = &H100000
'    CF_NOVECTORFONTS = &H800&
'    CF_NOVERTFONTS = &H1000000
'    CF_NULL = 0
'    CF_OEMTEXT = 7
'    CF_OWNERDISPLAY = &H80
'    CF_PALETTE = 9
'    CF_PENDATA = 10
'    CF_PRINTERFONTS = &H2
'    CF_PRIVATEFIRST = &H200
'    CF_PRIVATELAST = &H2FF
'    CF_REJECT = &H1
'    CF_RETEXTOBJ = ("RichEdit Text and Objects")
'    CF_RIFF = 11
'    CF_RTF = ("Rich Text Format")
'    CF_RTFNOOBJS = ("Rich Text Format Without Objects")
'    CF_SCALABLEONLY = &H20000
'    CF_SCREENFONTS = &H1
'    CF_SCRIPTSONLY = CF_ANSIONLY
'    CF_SELECTACTIVATEAS = &H10&
'    CF_SELECTCONVERTTO = &H8&
'    CF_SELECTSCRIPT = &H400000
'    CF_SETACTIVATEDEFAULT = &H4&
'    CF_SETCONVERTDEFAULT = &H2&
'    CF_SHOWHELP = &H4&
'    CF_SHOWHELPBUTTON = &H1&
'    CF_SYLK = 4
'    CF_TEXT = 1
'    CF_TIFF = 6
'    CF_TOPDOWN_DIB = 20
'    CF_TTONLY = &H40000
'    CF_UNICODETEXT = 13
'    CF_USESTYLE = &H80&
'    CF_WAVE = 12
'    CF_WYSIWYG = &H8000
'End Enum

Private Const NO_CB_OPEN_ERROR = 0
Private Const NO_CB_OPENED = 0
Private Const NO_CB_FORMAT_AVAILABLE = 0
Private Const NO_CB_VIWER = 0

Private Sub Command1_Click()
    CBSetText "This is a ClipBoard Test!!!"
End Sub

Private Sub Command2_Click()
    List1.AddItem CBGetText()
    List1.ListIndex = List1.ListCount - 1
End Sub


Private Sub Command3_Click()
    
    If GetOpenClipboardWindow() = NO_CB_OPENED Then
        If OpenClipboard(Me.hwnd) <> NO_CB_OPEN_ERROR Then
            EmptyClipboard
            CloseClipboard
        End If
    End If
    
End Sub

Private Sub Command4_Click()
    List1.Clear
End Sub

Private Sub Command5_Click()
    
    If GetOpenClipboardWindow() = NO_CB_OPENED Then
        If OpenClipboard(Me.hwnd) <> NO_CB_OPEN_ERROR Then
            EmptyClipboard
            CloseClipboard
        End If
    End If
    
End Sub

Private Sub Command6_Click()
    Set Image2.Picture = CBGetPicture()
    Image2.Refresh
End Sub

Private Sub Command7_Click()
    CBSetPicture Image1.Picture.handle
End Sub

Private Sub Command8_Click()
    Set Image2.Picture = Nothing
    Image2.Refresh
End Sub

Private Sub Form_Load()

'    Debug.Print RegisterClipboardFormat("GioRock Clipboard Format" + Chr$(0))
'    CloseClipboard
    
'    CBEnumerateFormats
    
    If GetOpenClipboardWindow() <> NO_CB_OPENED Then
        CloseClipboard
        SetClipboardViewer Me.hwnd
    End If
    
    sME(0) = "Clipboard open error!!!"
    sME(1) = "Not Clipboard BITMAP format available!!!"
    sME(2) = "Not Clipboard TEXT format available!!!"
    sME(3) = "Clipboard already opened by other application!!!"
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If GetOpenClipboardWindow() = Me.hwnd Then
'        EmptyClipboard
        CloseClipboard
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
    Set Form1 = Nothing
End Sub

Private Sub CBEnumerateFormats()
Dim fmt As Long, nfmt As Long
Dim s As String * 256

    If GetOpenClipboardWindow() = NO_CB_OPENED Then
        
        If OpenClipboard(Me.hwnd) <> NO_CB_OPEN_ERROR Then
        
            fmt = EnumClipboardFormats(&H0&)
        
            Do
                nfmt = EnumClipboardFormats(fmt)
                If nfmt = NO_CB_OPEN_ERROR Then
                    Exit Do
                Else
                    fmt = nfmt
                    s = String$(256, 0)
                    GetClipboardFormatName fmt, ByVal s, 256
                    List1.AddItem CStr(nfmt) + " = " + s + vbCrLf
                End If
            Loop
            
            CloseClipboard
            
            Exit Sub
            
        Else
            MsgError CB_OPEN_ERROR
        End If
        
    Else
        MsgError CB_ALREADY_OPEN
    End If
    
End Sub


Private Sub CBSetText(ByVal sCBText As String)
Dim hMem As Long, hPtr As Long, lLenBuffer As Long
Dim s As String
    
    If GetOpenClipboardWindow() = NO_CB_OPENED Then
        
        If OpenClipboard(Me.hwnd) <> NO_CB_OPEN_ERROR Then
        
            EmptyClipboard
            
            lLenBuffer = Len(sCBText) + 1
            s = String$(lLenBuffer, 0)
            Mid$(s, 1, lLenBuffer - 1) = sCBText
            
            hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_DDESHARE, lLenBuffer)
            hPtr = GlobalLock(hMem)
            
            MoveMemory ByVal hPtr, ByVal s, lLenBuffer
            
            GlobalUnlock hMem
            
            SetClipboardData CF_TEXT, hMem
            
            CloseClipboard
            
        Else
            MsgError CB_OPEN_ERROR
        End If
        
    Else
        MsgError CB_ALREADY_OPEN
    End If
    
'    GlobalUnlock hMem ' Never leave hMem Locked with SetClipboardData
'    GlobalFree hMem   ' Never Free hMem with SetClipboardData
                       ' Uses GMEM_MOVEABLE Or GMEM_DDESHARE
    
End Sub

Private Function CBGetText() As String
Dim hMem As Long, lLenBuffer As Long
Dim s As String
    
    If GetOpenClipboardWindow() = NO_CB_OPENED Then
        
        If OpenClipboard(Me.hwnd) <> NO_CB_OPEN_ERROR Then
        
            If IsClipboardFormatAvailable(CF_TEXT) <> NO_CB_FORMAT_AVAILABLE Then
            
                hMem = GetClipboardData(CF_TEXT)
                
                lLenBuffer = GlobalSize(hMem)
                
                s = String$(lLenBuffer, 0)
                
                MoveMemory ByVal s, ByVal hMem, lLenBuffer
                
                CloseClipboard
                
                If s = "" Then: Exit Function
                
                CBGetText = Left$(s, InStr(s, Chr$(0)) - 1)
                
                Exit Function
            
            Else
                CloseClipboard
                MsgError CB_NO_TEXT_FORMAT_AVAILABLE
            End If
            
        Else
            MsgError CB_OPEN_ERROR
        End If
        
    Else
        MsgError CB_ALREADY_OPEN
    End If
    
End Function

Private Sub CBSetPicture(ByVal hPic As Long)
Dim hMem As Long

    If GetOpenClipboardWindow() = NO_CB_OPENED Then
        
        If OpenClipboard(Me.hwnd) <> NO_CB_OPEN_ERROR Then
        
            EmptyClipboard
            
            hMem = CopyImage(hPic, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG Or LR_CREATEDIBSECTION)
            
            SetClipboardData CF_BITMAP, hMem
            
            CloseClipboard
            
        Else
            MsgError CB_OPEN_ERROR
        End If
        
    Else
        MsgError CB_ALREADY_OPEN
    End If

End Sub

Private Function CBGetPicture() As Object
Dim hMem As Long, R As Long
Dim IID_IDispatch As GUID
Dim OPic As OLEPIC, IObj As Object

    If GetOpenClipboardWindow() = NO_CB_OPENED Then

        If OpenClipboard(Me.hwnd) <> NO_CB_OPEN_ERROR Then
        
            If IsClipboardFormatAvailable(CF_BITMAP) <> NO_CB_FORMAT_AVAILABLE Then
            
                hMem = GetClipboardData(CF_BITMAP)
                
                With IID_IDispatch
                    .Data1 = &H20400
                    .Data4(0) = &HC0
                    .Data4(7) = &H46
                End With
    
                With OPic
                    .Size = Len(OPic)        'Lunghezza della struttura.
                    .tType = vbPicTypeBitmap 'Tipo dell'immagine (bitmap).
                    .hBmp = hMem             'L'handle dell'immagine.
'                    .hPal = hMem + 40       ' 40 Len BITMAP structure before palette
                End With
                
                R = OleCreatePictureIndirect(OPic, IID_IDispatch, 0, IObj)
                
                CloseClipboard
                
                Set CBGetPicture = IObj
                
                Set IObj = Nothing
                
                Exit Function
                
            Else
                CloseClipboard
                MsgError CB_NO_BITMAP_FORMAT_AVAILABLE
            End If
            
        Else
            MsgError CB_OPEN_ERROR
        End If
        
    Else
        MsgError CB_ALREADY_OPEN
    End If
    
    Set IObj = Nothing
    Set CBGetPicture = Nothing
    
End Function


Private Sub MsgError(eErr As eCBERRORMSG)
    MsgBox sME(eErr), vbInformation, App.EXEName
End Sub
