VERSION 5.00
Begin VB.UserControl xDirectory 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   KeyPreview      =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
End
Attribute VB_Name = "xDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'//---------------------------------------------------------------------------------------
'xDirectory Control
'//---------------------------------------------------------------------------------------
' Module    : xDirectory
' DateTime  : 06/10/2004 17:46
' Author    : Gary Noble
' Purpose   : Simluates A Directory Tree
' Assumes   : Nothing - Dependency Free
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
' Notes     : Thanks To CarlesPV For his Directory Tree Control
'             Thanks To Ulli For his SafeArray Code For Returning Array Parameters
'//---------------------------------------------------------------------------------------
Option Explicit

Private WithEvents m_cScrollBar As IAPP_ScrollBars
Attribute m_cScrollBar.VB_VarHelpID = -1
Private m_arrWidths() As String
Private m_lFirstSearch As Long

Public Enum eDisplayType
    edt_Fonts = 0
    edt_Drives = 1
    edt_Custom = 2
End Enum

Private m_lType As eDisplayType

Private Enum eMsg
    WM_SETREDRAW = &HB
End Enum

Private Type RECT
    Left                           As Long
    Top                            As Long
    Right                          As Long
    Bottom                         As Long
End Type

'Api declares
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Dim llastdrawn                                      As Long
Dim lfirstdrawn                                     As Long
Dim m_lSelectionID                                  As Long
Dim bHasFocus                                       As Boolean
Dim m_Button                                        As Integer
Dim m_sLongestString                                As String
Dim m_lLongestString                                As Long
Dim m_lLongestId                                    As Long

Private Const MaxDims As Long = 8

Private Type POINTAPI
    X                                               As Long
    Y                                               As Long
End Type


Private Const BITSPIXEL                             As Integer = 12
Private Const PS_SOLID                              As Integer = 0

'-- Colour functions:
Private Const OPAQUE                                As Integer = 2
Private Const TRANSPARENT                           As Integer = 1

Private Const DT_LEFT                               As Long = &H0
Private Const DT_RIGHT                              As Long = &H2
Private Const DT_END_ELLIPSIS                       As Long = &H8000&

'-- Scrolling and region functions:
Private Const LF_FACESIZE                           As Integer = 32
Private Type LOGFONT
    lfHeight                                        As Long
    lfWidth                                         As Long
    lfEscapement                                    As Long
    lfOrientation                                   As Long
    lfWeight                                        As Long
    lfItalic                                        As Byte
    lfUnderline                                     As Byte
    lfStrikeOut                                     As Byte
    lfCharSet                                       As Byte
    lfOutPrecision                                  As Byte
    lfClipPrecision                                 As Byte
    lfQuality                                       As Byte
    lfPitchAndFamily                                As Byte
    lfFaceName(LF_FACESIZE)                         As Byte
End Type
Private Const DST_ICON                              As Long = &H3
Private Const DSS_DISABLED                          As Long = &H20
Private Const DSS_MONO                              As Long = &H80
Private Const CLR_INVALID                           As Integer = -1
Private Type PictDesc
    cbSizeofStruct                                  As Long
    picType                                         As Long
    hImage                                          As Long
    xExt                                            As Long
    yExt                                            As Long
End Type

Private Type Guid
    Data1                                           As Long
    Data2                                           As Integer
    Data3                                           As Integer
    Data4(0 To 7)                                   As Byte
End Type
''------------------------------------------------------------------------------
'-- Image list Declares:
''------------------------------------------------------------------------------
'-- Create/Destroy functions:
Private Const ILC_MASK                              As Long = 1
Private Const ILC_COLOR32                           As Long = &H20&
'-- Modification/deletion functions:
'-- Image information functions:
Private Type IMAGEINFO
    hBitmapImage                                    As Long
    hBitmapMask                                     As Long
    cPlanes                                         As Long
    cBitsPerPixel                                   As Long
    rcImage                                         As RECT
End Type
'-- Create a new icon based on an image list icon:
'-- Merge and move functions:
Private Type IMAGELISTDRAWPARAMS
    cbSize                                          As Long
    hIml                                            As Long
    i                                               As Long
    hdcDst                                          As Long
    X                                               As Long
    Y                                               As Long
    cx                                              As Long
    cy                                              As Long
    xBitmap                                         As Long
    '--        // x offest from the upperleft of bitmap
    yBitmap                                         As Long
    '--        // y offset from the upperleft of bitmap
    rgbBk                                           As Long
    rgbFg                                           As Long
    fStyle                                          As Long
    dwRop                                           As Long
End Type
Private Const ILD_TRANSPARENT                       As Integer = 1
Private Const ILD_BLEND25                           As Integer = 2
Private Const ILD_SELECTED                          As Integer = 4
Private Const FORMAT_MESSAGE_FROM_SYSTEM            As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS         As Long = &H200
Private Type ICONINFO
    fIcon                                           As Long
    xHotspot                                        As Long
    yHotspot                                        As Long
    hbmMask                                         As Long
    hbmColor                                        As Long
End Type

Private m_lHdc                                      As Long
Private m_lHBmp                                     As Long
Private m_lHBmpOld                                  As Long
Private m_lhPalOld                                  As Long
Private m_pic                                       As StdPicture
Private m_sFileName                                 As String
Private m_lXOriginOffset                            As Long
Private m_lYOriginOffset                            As Long
Private m_lBitmapW                                  As Long
Private m_lBitmapH                                  As Long

Private Type BITMAP
    bmType                                          As Long
    bmWidth                                         As Long
    bmHeight                                        As Long
    bmWidthBytes                                    As Long
    bmPlanes                                        As Integer
    bmBitsPixel                                     As Integer
    bmBits                                          As Long
End Type

Private m_sCurrentSystemThemename                   As String

Private Type OSVERSIONINFO
    dwVersionInfoSize                               As Long
    dwMajorVersion                                  As Long
    dwMinorVersion                                  As Long
    dwBuildNumber                                   As Long
    dwPlatformId                                    As Long
    szCSDVersion(0 To 127)                          As Byte
End Type
Private Const IDC_HAND                              As Long = 32649
Private Const IDC_ARROW                             As Long = 32512
Private Const VER_PLATFORM_WIN32_NT                 As Integer = 2
Private Type TRIVERTEX
    X                                               As Long
    Y                                               As Long
    Red                                             As Integer
    Green                                           As Integer
    Blue                                            As Integer
    Alpha                                           As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft                                       As Long
    LowerRight                                      As Long
End Type
Private Type GRADIENT_TRIANGLE
    Vertex1                                         As Long
    Vertex2                                         As Long
    Vertex3                                         As Long
End Type
Public Enum GradientFillRectType
    GRADIENT_FILL_RECT_H = 0
    GRADIENT_FILL_RECT_V = 1
End Enum
#If False Then
    Private GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V
#End If


Private m_bHandCursor                               As Boolean
Private m_bIsXp                                     As Boolean
Public m_bIsNt                                      As Boolean
Private m_bIs2000OrAbove                            As Boolean
Private m_bHasGradientAndTransparency               As Boolean

Dim lLasty As Single
Dim llastx As Single
Private m_lptrVb6ImageList                          As Long
Private m_lIconWidth                                As Long
Private m_lIconHeight                               As Long
Private m_vImageList                                As Variant
Private m_lSelectedID As Long
Private m_lScrollbarItemCount As Long
Private bDBLClick As Boolean
Private m_HoverId As Long
Dim m_lLastIDDrawn As Long
Dim m_bDrawingSelectedNode As Boolean

Public Enum eNodeState
    ens_NotExpanded = 0
    ens_ExpandedOnce = 1
End Enum

Private Type tNodeData
    Level                                           As Long
    State                                           As eNodeState
    Caption                                         As String
    ItemData                                        As String
    Tag                                             As String
    Path                                            As String
    Expanded                                        As Boolean
    oRect                                           As RECT
    Children()                                      As Long
    Icon                                            As Long
    bSelected                                       As Boolean
    Parent                                          As Long
    ID                                              As Long
    HasChildren                                     As Boolean
End Type

Private Const MAX_PATH                              As Long = 260
Private Const MAXDWORD                              As Long = &HFFFF
Private Const INVALID_HANDLE_VALUE                  As Long = -1
Private Const FILE_ATTRIBUTE_ARCHIVE                As Long = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY              As Long = &H10
Private Const FILE_ATTRIBUTE_HIDDEN                 As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL                 As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY               As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM                 As Long = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY              As Long = &H100
Private Const FILE_ATTRIBUTE_RODIRECTORY            As Long = FILE_ATTRIBUTE_DIRECTORY + FILE_ATTRIBUTE_READONLY

Private Type FILETIME
    dwLowDateTime                                   As Long
    dwHighDateTime                                  As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes                                As Long
    ftCreationTime                                  As FILETIME
    ftLastAccessTime                                As FILETIME
    ftLastWriteTime                                 As FILETIME
    nFileSizeHigh                                   As Long
    nFileSizeLow                                    As Long
    dwReserved0                                     As Long
    dwReserved1                                     As Long
    cFileName                                       As String * MAX_PATH
    cAlternate                                      As String * 14
End Type

Private Const CSIDL_DESKTOP                         As Long = &H0
Private Const CSIDL_INTERNET                        As Long = &H1
Private Const CSIDL_PROGRAMS                        As Long = &H2
Private Const CSIDL_CONTROLS                        As Long = &H3
Private Const CSIDL_PRINTERS                        As Long = &H4
Private Const CSIDL_PERSONAL                        As Long = &H5
Private Const CSIDL_FAVORITES                       As Long = &H6
Private Const CSIDL_STARTUP                         As Long = &H7
Private Const CSIDL_RECENT                          As Long = &H8
Private Const CSIDL_SENDTO                          As Long = &H9
Private Const CSIDL_BITBUCKET                       As Long = &HA
Private Const CSIDL_STARTMENU                       As Long = &HB
Private Const CSIDL_DESKTOPDIRECTORY                As Long = &H10
Private Const CSIDL_DRIVES                          As Long = &H11
Private Const CSIDL_NETWORK                         As Long = &H12
Private Const CSIDL_NETHOOD                         As Long = &H13
Private Const CSIDL_FONTS                           As Long = &H14
Private Const CSIDL_TEMPLATES                       As Long = &H15
Private Const CSIDL_COMMON_STARTMENU                As Long = &H16
Private Const CSIDL_COMMON_PROGRAMS                 As Long = &H17
Private Const CSIDL_COMMON_STARTUP                  As Long = &H18
Private Const CSIDL_COMMON_DESKTOPDIRECTORY         As Long = &H19
Private Const CSIDL_APPDATA                         As Long = &H1A
Private Const CSIDL_PRINTHOOD                       As Long = &H1B
Private Const CSIDL_ALTSTARTUP                      As Long = &H1D
Private Const CSIDL_COMMON_ALTSTARTUP               As Long = &H1E
Private Const CSIDL_COMMON_FAVORITES                As Long = &H1F
Private Const CSIDL_INTERNET_CACHE                  As Long = &H20
Private Const CSIDL_COOKIES                         As Long = &H21
Private Const CSIDL_HISTORY                         As Long = &H22

Private Const SHGFI_LARGEICON                       As Long = &H0
Private Const SHGFI_SMALLICON                       As Long = &H1
Private Const SHGFI_OPENICON                        As Long = &H2
Private Const SHGFI_SHELLICONSIZE                   As Long = &H4
Private Const SHGFI_PIDL                            As Long = &H8
Private Const SHGFI_USEFILEATTRIBUTES               As Long = &H10
Private Const SHGFI_ICON                            As Long = &H100
Private Const SHGFI_DISPLAYNAME                     As Long = &H200
Private Const SHGFI_TYPENAME                        As Long = &H400
Private Const SHGFI_ATTRIBUTES                      As Long = &H800
Private Const SHGFI_ICONLOCATION                    As Long = &H1000
Private Const SHGFI_EXETYPE                         As Long = &H2000
Private Const SHGFI_SYSICONINDEX                    As Long = &H4000
Private Const SHGFI_LINKOVERLAY                     As Long = &H8000
Private Const SHGFI_SELECTED                        As Long = &H10000
Private Const SHGFI_ATTR_SPECIFIED                  As Long = &H20000

Private Const S_OK                                  As Long = &H0

Private Const SI_FOLDER_CLOSED                      As Long = &H3
Private Const SI_FOLDER_OPEN                        As Long = &H4

Private Type SHFILEINFO
    hIcon                                           As Long
    iIcon                                           As Long
    dwAttributes                                    As Long
    szDisplayName                                   As String * MAX_PATH
    szTypeName                                      As String * 80
End Type


Private Enum ArrayFeatures
    FADF_AUTO = &H1    'Array is allocated on the stack
    FADF_STATIC = &H2    'Array is statically allocated
    FADF_EMBEDDED = &H4    'Array is embedded in a structure
    FADF_FIXEDSIZE = &H10    'Array may not be resized or reallocated
    FADF_BSTR = &H100    'An array of BSTRs
    FADF_UNKNOWN = &H200    'An array of IUnknown*
    FADF_DISPATCH = &H400    'An array of IDispatch*
    FADF_VARIANT = &H800    'An array of VARIANTs
    FADF_RESERVED = &HFFFFF0E8    'Bits reserved for future use
End Enum
#If False Then    'keep capitalization
    Private FADF_AUTO, FADF_STATIC, FADF_EMBEDDED, FADF_FIXEDSIZE, FADF_BSTR, FADF_UNKNOWN, FADF_DISPATCH, FADF_VARIANT, FADF_RESERVED
#End If

Private Type SAFEARRAYBOUND
    NumElements                                     As Long
    LBound                                          As Long
    UBound                                          As Long
End Type

Private Type SAFEARRAYDESCRIPTOR
    NumDims                                         As Integer    'number of dimensions
    Features                                        As Integer    'feature bits
    ElementSize                                     As Long    'size of one element
    Locks                                           As Long    'number of locks
    PtrToData                                       As Long    'pointer to first element
    Bounds(1 To MaxDims)                            As SAFEARRAYBOUND    'number of elements and lbound/ubound for each dimension
End Type

Private m_bInitialized                              As Boolean
Private m_Nodes()                                   As tNodeData
Private m_DisplayNodes()                            As Long
Private m_HoverSelection                            As Boolean
Private Const m_def_HoverSelection                  As Boolean = False

Private m_BackGroundPicture                         As Picture

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
        ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
        ByVal nCount As Long, _
        lpObject As Any) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
        ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long
Private Declare Function DrawEdge Lib "user32" _
        (ByVal hDC As Long, qrc As RECT, ByVal edge As _
        Long, ByVal grfFlags As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
        ByVal nWidth As Long, _
        ByVal crColor As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        lpPoint As POINTAPI) As Long

Private Declare Function DrawEdgeAPI Lib "user32" Alias "DrawEdge" (ByVal hDC As Long, _
        qrc As RECT, _
        ByVal edge As Long, _
        ByVal grfFlags As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, _
        ByVal crColor As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, _
        ByVal crColor As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, _
        ByVal nBkMode As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, _
        ByVal lpszExeFileName As String, _
        ByVal nIconIndex As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, _
        ByVal lpsz As String, _
        ByVal un1 As Long, _
        ByVal n1 As Long, _
        ByVal n2 As Long, _
        ByVal un2 As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
        ByVal ptX As Long, _
        ByVal ptY As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, _
        lpRect As RECT, _
        ByVal hBrush As Long) As Long

Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
        ByVal xLeft As Long, _
        ByVal yTop As Long, _
        ByVal hIcon As Long, _
        ByVal cxWidth As Long, _
        ByVal cyWidth As Long, _
        ByVal istepIfAniCur As Long, _
        ByVal hbrFlickerFreeDraw As Long, _
        ByVal diFlags As Long) As Boolean

Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, _
        ByVal X1 As Long, _
        ByVal y1 As Long, _
        ByVal x2 As Long, _
        ByVal y2 As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, _
        ByVal hBrush As Long, _
        ByVal lpDrawStateProc As Long, _
        ByVal lParam As Long, _
        ByVal wParam As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal fuFlags As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
        ByVal HPALETTE As Long, _
        pccolorref As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
        pSrc As Any, _
        ByVal ByteLen As Long)

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, _
        lpSource As Any, _
        ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long, _
        Arguments As Long) As Long

Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
        ByVal nCount As Long, _
        lpObject As Any) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal fStyle As Long) As Long

Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal diIgnore As Long) As Long

Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal hIml As Long) As Long

Private Declare Function ImageList_GetImageRect Lib "comctl32.dll" (ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT) As Long

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
        lpRect As RECT) As Long

Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, _
        ByVal pszClassList As Long) As Long

Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, _
        ByVal dwMaxNameChars As Long, _
        ByVal pszColorBuff As Long, _
        ByVal cchMaxColorChars As Long, _
        ByVal pszSizeBuff As Long, _
        ByVal cchMaxSizeChars As Long) As Long

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, _
        ByVal lpCursorName As Long) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
        lpDeviceName As Any, _
        lpOutput As Any, _
        lpInitData As Any) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, _
        ByVal lpStr As String, _
        ByVal nCount As Long, _
        lpRect As RECT, _
        ByVal wFormat As Long) As Long

Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, _
        ByVal lpStr As Long, _
        ByVal nCount As Long, _
        lpRect As RECT, _
        ByVal wFormat As Long) As Long


Private Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hDC As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long

Private Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, _
        pVertex As TRIVERTEX, _
        ByVal dwNumVertex As Long, _
        pMesh As GRADIENT_RECT, _
        ByVal dwNumMesh As Long, _
        ByVal dwMode As Long) As Long

Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, _
        ByVal lHDC As Long, _
        ByVal iPartId As Long, _
        ByVal iStateId As Long, _
        pRect As RECT, _
        pClipRect As RECT) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
        lpPoint As POINTAPI) As Long

Private Declare Function IspbAppThemed Lib "uxtheme.dll" () As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef PIDL As Long) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal PIDL As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)


Private Declare Function hTable Lib "msvbvm50.dll" Alias "VarPtr" (Table() As Any) As Long

Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As _
        Long, ByVal wParam As Long, lParam As Any) As Long


'-- Events
Event SelectionChanged(ByVal sValue As String, ByVal sTag As String)
Event RightClick(ByVal sValue As String)
Event LocatingComplete()
Event Locating(ByVal sDirectory As String)
'Event Declarations:
Event NodeDblClick(ByVal ID As Long, ByVal Caption As String, ByVal sTag As String)
Event BeforeCustomNodeExpand(lNode As Long, Caption As String, State As Boolean)


'//---------------------------------------------------------------------------------------
' Procedure : pvGetArrayDescriptor
' Type      : Function
' DateTime  : 06/10/2004 17:23
' Author    : Gary Noble
' Purpose   :
' Returns   : SAFEARRAYDESCRIPTOR
' Notes     : Taken From Ullis code. All Credit To Him!
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Function pvGetArrayDescriptor(ByVal hTable As Long) As SAFEARRAYDESCRIPTOR

'param hTable must point to a pointer which in turn points to the array descriptor
'
'you get the hTable parameter by calling hTable(your_table_name)
'hTable is in fact a disguise for the VarPtr function (which unfortunately does not
'accept tables() )
'
'so the function call for this function should look like this:
'
'pvGetArrayDescriptor(hTable(your_table_name)) 'returns SAFEARRAYDESCRIPTOR for your_table_name
'
'one little drawback though:
'apparently VB does not store variable (redimmable) tables of variable length strings
'as safearrays, so this set of routines does not work with this kind of tables.
'
'it's okay however with fixed (non-redimmable) tables of variable length strings
'and with variable (redimmable) tables of fixed length strings

    Dim PtrToDesc As Long
    Dim i         As Long

    CopyMemory PtrToDesc, ByVal hTable, 4
    If PtrToDesc Then
        With pvGetArrayDescriptor
            CopyMemory .NumDims, ByVal PtrToDesc, 16    'get the first 16 bytes (NumDims..PtrToData)
            If .NumDims <= MaxDims Then    'to prevent out of range indexing
                PtrToDesc = PtrToDesc + 16    'adjust pointer
                For i = .NumDims To 1 Step -1    'in reverse order; the m.s. dimension is rightmost
                    CopyMemory .Bounds(i), ByVal PtrToDesc, 8    'get Number of Elements and LBound
                    PtrToDesc = PtrToDesc + 8    'adjust pointer
                    With .Bounds(i)
                        .UBound = .LBound + .NumElements - 1    'calculate UBound
                    End With    '.BOUNDS(I)
                Next i
            End If
        End With    'pvGetArrayDescriptor
    End If

CleanExit:

    On Error GoTo 0
    Exit Function

End Function

'//---------------------------------------------------------------------------------------
' Procedure : pbIsDimmed
' Type      : Function
' DateTime  : 06/10/2004 17:24
' Author    : Gary Noble
' Purpose   : Taken from Ullis Code , Returns The number Of Dims In An Array
' Returns   : Boolean
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Function pbIsDimmed(ByVal hTable As Long) As Boolean

'hTable points to a pointer which in turn points to the array desriptor

'you get the hTable parameter by calling hTable(your_table_name) so the
'function call should look like this:

'If pbIsDimmed(hTable(your_table_name)) Then ...

    pbIsDimmed = pvGetArrayDescriptor(hTable).NumDims

End Function



'//---------------------------------------------------------------------------------------
' Procedure : plAddNode
' Type      : Function
' DateTime  : 06/10/2004 17:24
' Author    : Gary Noble
' Purpose   : Adds A Node To The Node Array
' Returns   : Long
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Function plAddNode(Caption As String, Path As String, Icon As Long, Parent As Long) As Long

    Dim oDescriptor As SAFEARRAYDESCRIPTOR
    Dim lID As Long
    Dim lIDParent As Long
    Dim lLevel As Long


    '-- Add A Reference To The Parent Node If Any
    If Parent > 0 Then

        If pbIsDimmed(hTable(m_Nodes(Parent).Children)) Then
            ReDim Preserve m_Nodes(Parent).Children(UBound(m_Nodes(Parent).Children) + 1)
        Else
            ReDim m_Nodes(Parent).Children(0)
        End If

        '-- Increment The Node Level
        lLevel = m_Nodes(Parent).Level + 1

    End If

    '-- Rebound Our Nodes Array
    If pbIsDimmed(hTable(m_Nodes)) Then
        ReDim Preserve m_Nodes(UBound(m_Nodes) + 1) As tNodeData
    Else
        ReDim m_Nodes(1) As tNodeData
    End If

    '-- Set The Node Data
    With m_Nodes(UBound(m_Nodes))
        .Caption = Caption
        .State = ens_NotExpanded
        .Icon = Icon
        .Path = Path
        .ID = lID
        .Level = lLevel
    End With

    '-- Just Incase
    If lLevel > 0 Then
        '-- Parent Id Locator
        m_Nodes(Parent).Children(lIDParent) = lID
        m_Nodes(UBound(m_Nodes)).Parent = Parent
    End If

    '-- Return The node
    plAddNode = UBound(m_Nodes)

End Function


'//---------------------------------------------------------------------------------------
' Procedure : Init
' Type      : Function
' DateTime  : 06/10/2004 17:22
' Author    : Gary Noble
' Purpose   : Initial Load Call
' Returns   : Variant
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Public Function Init(Optional ByVal sPath As String, Optional lDisplayType As eDisplayType = edt_Drives)


'-- Clean And Set Defaults
    m_bDrawingSelectedNode = False
    m_lSelectedID = -1

    m_lType = lDisplayType

    '-- Clear Or Nodes
    Erase m_Nodes
    Erase m_DisplayNodes

    '-- Redraw
    pvRedraw
    DoEvents
    UserControl.MousePointer = vbHourglass



    If lDisplayType = edt_Drives Then

        '-- Add The Drive List
        pvAddDrives

        '-- Create The Display nodes
        pvMakeDisplayNodes

        '-- Redraw
        pvRedraw

        '-- Find The start Path If Any
        If Len(Trim(sPath)) > 0 Then
            pvDisplayStartupPath sPath
        Else
            m_lSelectionID = 1
            m_lSelectedID = m_DisplayNodes(1)
        End If

    ElseIf lDisplayType = edt_Fonts Then
        GetFonts
    ElseIf lDisplayType = edt_Custom Then
        'Code Here
    End If

    pvMakeDisplayNodes
    UserControl.MousePointer = vbDefault



End Function

'//---------------------------------------------------------------------------------------
' Procedure : Clear
' Type      : Sub
' DateTime  : 06/10/2004 17:23
' Author    : Gary Noble
' Purpose   : Clear
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Public Sub Clear()

    Cls
    '-- Hide The ScrollBars
    m_cScrollBar.Visible(efsHorizontal) = False
    m_cScrollBar.Visible(efsVertical) = False

    Erase m_DisplayNodes
    Erase m_Nodes
    pvMakeDisplayNodes
    pvMakeScrollScrollbarVisible
    pvRedraw

End Sub
'//---------------------------------------------------------------------------------------
' Procedure : pvMakeDisplayNodes
' Type      : Sub
' DateTime  : 06/10/2004 17:26
' Author    : Gary Noble
' Purpose   : Creats An Array Of displed Nodes Used For Drawing To The Screen
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvMakeDisplayNodes()

    Dim i As Long
    Dim lcount As Long


    Erase m_DisplayNodes

    '-- Bail
    If Not pbIsDimmed(hTable(m_Nodes)) Then Exit Sub

    '-- Scrollbar Count
    m_lScrollbarItemCount = 0
    m_lLongestString = m_lSelectedID

    lcount = UBound(m_Nodes)

    If pbIsDimmed(hTable(m_Nodes)) Then

        '-- Loop Through The Nodes And Add It To The Node Array, Only If the Level Is Top Level
        For i = 0 To lcount

            If m_Nodes(i).Level = 0 Then

                '-- Increment The Scrollbar
                m_lScrollbarItemCount = m_lScrollbarItemCount + 1
                If pbIsDimmed(hTable(m_DisplayNodes)) Then
                    ReDim Preserve m_DisplayNodes(UBound(m_DisplayNodes) + 1)
                Else
                    ReDim m_DisplayNodes(0)
                End If

                m_DisplayNodes(UBound(m_DisplayNodes)) = i
                '-- get The Child Nodes
                If m_Nodes(i).Expanded Then pvMakeDisplayNodesX i
            End If

        Next

    End If

    pvMakeScrollScrollbarVisible
    pvRedraw

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvMakeDisplayNodesX
' Type      : Sub
' DateTime  : 06/10/2004 17:28
' Author    : Gary Noble
' Purpose   : The Same As pvMakeDisplayNodes Except For Children
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvMakeDisplayNodesX(lID As Long)
    Dim i As Long
    Dim lcount As Long

    If Not pbIsDimmed(hTable(m_Nodes)) Then Exit Sub


    lcount = UBound(m_Nodes)


    If Not pbIsDimmed(hTable(m_Nodes)) Then Exit Sub


    lcount = UBound(m_Nodes)

    If pbIsDimmed(hTable(m_Nodes)) Then

        For i = lID + 1 To lcount

            If m_Nodes(i).Level > 0 And m_Nodes(i).Parent = lID Then

                m_lScrollbarItemCount = m_lScrollbarItemCount + 1

                ReDim Preserve m_DisplayNodes(UBound(m_DisplayNodes) + 1)

                m_DisplayNodes(UBound(m_DisplayNodes)) = i

                If pbIsDimmed(hTable(m_Nodes(i).Children)) And m_Nodes(i).Expanded Then pvMakeDisplayNodesX i

            End If

        Next

    End If



End Sub


'//---------------------------------------------------------------------------------------
' Procedure : pvMakeScrollScrollbarVisible
' Type      : Sub
' DateTime  : 06/10/2004 17:28
' Author    : Gary Noble
' Purpose   : Hides Or Displays The Scrollbar
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvMakeScrollScrollbarVisible()

    Dim lcount As Long

    If Not pbIsDimmed(hTable(m_DisplayNodes)) Then GoTo CleanExit



CleanExit:

    Dim lVis As Long


    Dim i As Long

    '-- Horizontal Scrollbar
    For i = 1 To UBound(m_DisplayNodes)
        If TextWidth(Space(m_Nodes(i).Level * 6)) + TextWidth(m_Nodes(i).ItemData & "" & m_Nodes(i).Caption) > lVis Then
            lVis = TextWidth(Space(m_Nodes(i).Level * 6)) + TextWidth(m_Nodes(i).ItemData & "" & m_Nodes(i).Caption)
        End If
    Next

    m_cScrollBar.Visible(efsHorizontal) = lVis + 15 > ScaleWidth

    If m_cScrollBar.Visible(efsHorizontal) Then
        m_cScrollBar.SmallChange(efsHorizontal) = 20
        m_cScrollBar.LargeChange(efsHorizontal) = ScaleWidth / 3
        m_cScrollBar.Max(efsHorizontal) = lVis + 60 - ScaleWidth
    Else
        m_cScrollBar.Value(efsHorizontal) = 0
    End If


    '-- Vertical Scrollbar
    If (m_lScrollbarItemCount) - (ScaleHeight \ (TextHeight("Q,`") + 6)) - 1 > 0 Then
        m_cScrollBar.Max(efsVertical) = (m_lScrollbarItemCount) - (ScaleHeight \ (TextHeight("Q,`") + 6)) - 1
        m_cScrollBar.Visible(efsVertical) = True
        m_cScrollBar.SmallChange(efsVertical) = 1
        m_cScrollBar.LargeChange(efsVertical) = (llastdrawn - lfirstdrawn)
    Else
        m_cScrollBar.Visible(efsVertical) = False
        m_cScrollBar.Value(efsVertical) = 0
    End If

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvRedraw
' Type      : Sub
' DateTime  : 06/10/2004 17:28
' Author    : Gary Noble
' Purpose   : Draws The Actual Control
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvRedraw()

    On Error Resume Next

    Dim i As Long
    Dim rGlymph As RECT
    Dim xY As Long
    Dim rcFocus As RECT
    Dim fntOrig As String
    Dim fntOrigSize As Long

    fntOrig = UserControl.Font.Name
    fntOrigSize = UserControl.Font.Size

    '-- Paint The Background
    If Not Me.BackGroundPicture Is Nothing Then
        '-- We Don't Need To Clear The Screen If The Picture Is Set
        pvSetPicture Me.BackGroundPicture
        pvTileArea hDC, 0, 0, ScaleWidth, ScaleHeight
    Else
        Cls
    End If



    '-- Top CoOrdinate
    xY = 1

    '-- First Drawn Flag - Used To Get The Count Of Nodes Painted On The Screen
    lfirstdrawn = m_cScrollBar.Value(efsVertical) + 1
    llastdrawn = lfirstdrawn
    '-- Bail
    If Not pbIsDimmed(hTable(m_Nodes)) Then Exit Sub

    '-- Paint the Nodes
    For i = m_cScrollBar.Value(efsVertical) + 1 To UBound(m_DisplayNodes)

        If i > 1 Then If m_lType = edt_Fonts Then UserControl.FontName = m_Nodes(m_DisplayNodes(i)).Caption

        '-- Set The Working Rectangles
        If m_lType = edt_Drives Then
            m_Nodes(m_DisplayNodes(i)).oRect.Left = -m_cScrollBar.Value(efsHorizontal) + IIf(i = 1, 35, 35) + TextWidth(Space(m_Nodes(m_DisplayNodes(i)).Level * 6))
        ElseIf m_lType = edt_Fonts Then
            m_Nodes(m_DisplayNodes(i)).oRect.Left = -m_cScrollBar.Value(efsHorizontal) + 15
        ElseIf m_lType = edt_Custom Then
            m_Nodes(m_DisplayNodes(i)).oRect.Left = -m_cScrollBar.Value(efsHorizontal) + 15
        End If

        m_Nodes(m_DisplayNodes(i)).oRect.Top = xY
        m_Nodes(m_DisplayNodes(i)).oRect.Bottom = xY + TextHeight("Q") + 5    ' ScaleHeight
        m_Nodes(m_DisplayNodes(i)).oRect.Right = m_Nodes(m_DisplayNodes(i)).oRect.Left + 35 + 8 + (TextWidth(m_Nodes(m_DisplayNodes(i)).Caption))
        If m_lSelectedID = m_DisplayNodes(i) Then
            '-- Drawing Selected Node Flag
            m_bDrawingSelectedNode = True

            m_lSelectedID = m_DisplayNodes(i)
            m_lSelectionID = i
            If m_lSelectedID > 0 Then

                If bHasFocus Then
                    '-- Draw The Background
                    pvDrawBackground hDC, vbHighlight, vbHighlight, m_Nodes(m_DisplayNodes(i)).oRect.Left + 4, xY - 1, 6 + (TextWidth(m_Nodes(m_DisplayNodes(i)).Caption)), TextHeight("Q") + 3, False

                    '-- Draw The Focus Rect
                    LSet rcFocus = m_Nodes(m_DisplayNodes(i)).oRect
                    rcFocus.Left = (rcFocus.Left + 4)
                    rcFocus.Top = xY - 1
                    rcFocus.Right = rcFocus.Left + 6 + (TextWidth(m_Nodes(m_DisplayNodes(i)).Caption))
                    rcFocus.Bottom = rcFocus.Bottom - 3
                    DrawFocusRect hDC, rcFocus
                Else
                    pvDrawBackground hDC, vbButtonFace, vbButtonFace, m_Nodes(m_DisplayNodes(i)).oRect.Left + 4, xY - 1, 6 + (TextWidth(m_Nodes(m_DisplayNodes(i)).Caption)), TextHeight("Q") + 3, False
                End If

            End If

        End If

        LSet rGlymph = m_Nodes(m_DisplayNodes(i)).oRect
        rGlymph.Top = (xY) + (TextHeight("Q") \ 2) - (TextHeight("Q") \ 2)

        '-- Not used
        If m_HoverId = m_DisplayNodes(i) Then UserControl.Font.Underline = True


        '-- Paint the Node Text
        If m_Nodes(m_DisplayNodes(i)).HasChildren Then
            If m_lSelectedID = m_DisplayNodes(i) Then
                If m_lSelectedID > 0 Then
                    pvDrawText hDC, m_Nodes(m_DisplayNodes(i)).Caption, rGlymph.Left + 6, rGlymph.Top, rGlymph.Right, ScaleHeight, True, plTranslateColor(&H8000000E), False
                End If
            Else
                pvDrawText hDC, m_Nodes(m_DisplayNodes(i)).Caption, rGlymph.Left + 6, rGlymph.Top, rGlymph.Right, ScaleHeight, True, UserControl.ForeColor, False
            End If
        Else
            If m_lSelectedID = m_DisplayNodes(i) Then
                If m_lSelectedID > 0 Then
                    pvDrawText hDC, m_Nodes(m_DisplayNodes(i)).Caption, rGlymph.Left + 6, rGlymph.Top, rGlymph.Right, ScaleHeight, True, plTranslateColor(&H8000000E), False
                End If
            Else
                pvDrawText hDC, m_Nodes(m_DisplayNodes(i)).Caption, rGlymph.Left + 6, rGlymph.Top, rGlymph.Right, ScaleHeight, True, UserControl.ForeColor, False
            End If
        End If


        If m_HoverId = m_DisplayNodes(i) Then UserControl.Font.Underline = False

        '-- Paint The ItemData If Any
        If Len(RTrim(m_Nodes(m_DisplayNodes(i)).ItemData)) > 0 Then
            pvDrawText hDC, m_Nodes(m_DisplayNodes(i)).ItemData, TextWidth(m_Nodes(m_DisplayNodes(i)).Caption) + m_Nodes(m_DisplayNodes(i)).oRect.Left + 15, rGlymph.Top, rGlymph.Right + 500, ScaleHeight, True, IIf(m_Nodes(m_DisplayNodes(i)).ItemData = "(Nothing)", vbRed, vbBlue), False
        End If

        '-- Draw the Expand/Collapse Button
        If m_Nodes(m_DisplayNodes(i)).HasChildren Then
            If m_lType = edt_Drives Then
                rGlymph.Left = rGlymph.Left - 35
            ElseIf m_lType = edt_Fonts Then
                rGlymph.Left = rGlymph.Left - 15
            ElseIf m_lType = edt_Custom Then
                rGlymph.Left = rGlymph.Left - 15
            End If


            rGlymph.Top = (xY) + (TextHeight("Q") \ 2) - 7
            pvDrawOpenCloseGlyph UserControl.hwnd, UserControl.hDC, rGlymph, Not m_Nodes(m_DisplayNodes(i)).Expanded
        End If

        '-- Draw the Icon
        LSet rGlymph = m_Nodes(m_DisplayNodes(i)).oRect
        rGlymph.Left = rGlymph.Left - 19
        rGlymph.Top = (xY) + (TextHeight("Q") \ 2) - 7
        If m_lType = edt_Drives Then pvImageListDrawIcon m_lptrVb6ImageList, hDC, m_Nodes(m_DisplayNodes(i)).Icon, m_Nodes(m_DisplayNodes(i)).Icon, rGlymph.Left + 3, rGlymph.Top, False, False, False


        '-- Ofset The Original Rects
        xY = xY + TextHeight("Q") + 5
        If m_lType = edt_Drives Then
            m_Nodes(m_DisplayNodes(i)).oRect.Left = IIf(m_Nodes(m_DisplayNodes(i)).HasChildren, m_Nodes(m_DisplayNodes(i)).oRect.Left - 35, m_Nodes(m_DisplayNodes(i)).oRect.Left - 28)
        ElseIf m_lType = edt_Fonts Then
            m_Nodes(m_DisplayNodes(i)).oRect.Left = IIf(m_Nodes(m_DisplayNodes(i)).HasChildren, m_Nodes(m_DisplayNodes(i)).oRect.Left - 15, m_Nodes(m_DisplayNodes(i)).oRect.Left - 28)
        ElseIf m_lType = edt_Custom Then
            m_Nodes(m_DisplayNodes(i)).oRect.Left = IIf(m_Nodes(m_DisplayNodes(i)).HasChildren, m_Nodes(m_DisplayNodes(i)).oRect.Left - 15, m_Nodes(m_DisplayNodes(i)).oRect.Left - 28)
        End If
        m_Nodes(m_DisplayNodes(i)).oRect.Right = m_Nodes(m_DisplayNodes(i)).oRect.Left + 35 + 8 + (TextWidth(m_Nodes(m_DisplayNodes(i)).Caption))
        m_lLastIDDrawn = i
        llastdrawn = i

        If m_lType = edt_Fonts Then UserControl.Font.Bold = False: _
                UserControl.Font.Italic = False: UserControl.Font.Underline = False: _
                UserControl.Font.Size = fntOrigSize

        If xY > ScaleHeight Then Exit For

    Next

    If m_lType = edt_Fonts Then UserControl.FontName = fntOrig
    If m_lType = edt_Fonts Then UserControl.Font.Bold = False: UserControl.Font.Italic = False: _
            UserControl.Font.Underline = False: UserControl.Font.Size = fntOrigSize


End Sub


Private Sub m_cScrollBar_Change(eBar As EFSScrollBarConstants)
    DoEvents
    pvRedraw
End Sub

Private Sub m_cScrollBar_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
    DoEvents
    pvRedraw
End Sub

Private Sub m_cScrollBar_Scroll(eBar As EFSScrollBarConstants)
    DoEvents
    pvRedraw
End Sub

Private Sub m_cScrollBar_ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)
    DoEvents
    pvRedraw
End Sub

Private Sub UserControl_DblClick()
    bDBLClick = True
    UserControl_MouseDown m_Button, 0, llastx, lLasty
    bDBLClick = False

End Sub

Private Sub UserControl_EnterFocus()
    bHasFocus = True
End Sub

Private Sub UserControl_ExitFocus()
    bHasFocus = False
    pvRedraw
End Sub

Private Sub UserControl_GotFocus()
    bHasFocus = True
    pvRedraw
End Sub

Private Sub UserControl_Initialize()

    InitCommonControls
    pvVerInitialise
    pvGetSystemImageList SHGFI_SMALLICON

End Sub

Private Sub UserControl_LostFocus()
    bHasFocus = False
    pvRedraw
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lID As Long
    Dim sCap As String
    Dim bExpanded As Boolean

    On Error Resume Next

    Dim lIID As Long
    Dim bChevron As Boolean
    pvHittest X, Y, lIID, bChevron
    bHasFocus = True
    m_Button = Button
    lID = lIID

    If Button = vbLeftButton Then

        If lID - 1 > UBound(m_DisplayNodes) Then
            Exit Sub
        Else
            If lID = 0 Then Exit Sub

            '-- Only Expand the Node - Don't Select IT , Populate The Dirs
            If bChevron Then

                If m_lType = edt_Drives Then

                    If Not m_Nodes(m_DisplayNodes(lID)).State = ens_ExpandedOnce Then

                        If m_lType = edt_Drives Then

                            RaiseEvent Locating(m_Nodes(m_DisplayNodes(lID)).Path)

                            UserControl.MousePointer = vbHourglass
                            DoEvents
                            m_Nodes(m_DisplayNodes(lID)).Expanded = True

                            '-- Check To See If The Node Has Been Expanded Before
                            If m_Nodes(m_DisplayNodes(lID)).State = ens_NotExpanded Then
                                pvAddFolders m_DisplayNodes(lID)

                                '-- Set The Expanded Flag
                                m_Nodes(m_DisplayNodes(lID)).State = ens_ExpandedOnce
                                UserControl.MousePointer = vbDefault
                                m_Nodes(m_DisplayNodes(lID)).Expanded = True

                                '-- Make the DisplayNodes
                                pvMakeDisplayNodes

                                If Not pbIsDimmed(hTable(m_Nodes(m_DisplayNodes(lID)).Children)) Then
                                    m_Nodes(m_DisplayNodes(lID)).HasChildren = False
                                Else
                                    m_Nodes(m_DisplayNodes(lID)).ItemData = "(" & UBound(m_Nodes(m_DisplayNodes(lID)).Children) + 1 & ")"
                                End If
                            End If

                            RaiseEvent LocatingComplete
                        Else
                            If Not m_Nodes(m_DisplayNodes(lID)).Expanded = True Then
                                RaiseEvent BeforeCustomNodeExpand(lID, m_Nodes(m_DisplayNodes(lID)).Caption, Not m_Nodes(m_DisplayNodes(lID)).Expanded)
                            End If
                        End If
                    Else
                        m_Nodes(m_DisplayNodes(lID)).Expanded = Not m_Nodes(m_DisplayNodes(lID)).Expanded

                    End If
                    '-- Make the DisplayNodes

                    pvRedraw

                Else

                    m_Nodes(m_DisplayNodes(lID)).Expanded = Not m_Nodes(m_DisplayNodes(lID)).Expanded

                    If m_lType = edt_Custom Then
                        If m_Nodes(m_DisplayNodes(lID)).HasChildren Then
                            If Not m_Nodes(m_DisplayNodes(lID)).State = ens_ExpandedOnce Then
                                bExpanded = Not m_Nodes(m_DisplayNodes(lID)).Expanded
                                sCap = m_Nodes(m_DisplayNodes(lID)).Caption
                                RaiseEvent BeforeCustomNodeExpand(lID, sCap, bExpanded)
                                m_Nodes(m_DisplayNodes(lID)).State = ens_ExpandedOnce
                            End If
                        End If
                    End If

                End If
            Else
                m_lSelectedID = m_DisplayNodes(lID)

                RaiseEvent SelectionChanged(m_Nodes(m_DisplayNodes(lID)).Caption, m_Nodes(m_DisplayNodes(lID)).Path)


                '-- Expand And Select The Node , Populate The Dirs
                If bDBLClick Then

                    RaiseEvent NodeDblClick(lID, m_Nodes(m_DisplayNodes(lID)).Caption, m_Nodes(m_DisplayNodes(lID)).Path)

                    If m_lType = edt_Drives Then

                        '-- Check To See If The Node Has Been Expanded Before
                        If Not m_Nodes(m_DisplayNodes(lID)).State = ens_ExpandedOnce Then
                            RaiseEvent Locating(m_Nodes(m_DisplayNodes(lID)).Path)
                            UserControl.MousePointer = vbHourglass
                            DoEvents
                            RaiseEvent SelectionChanged(m_Nodes(m_DisplayNodes(lID)).Caption, m_Nodes(m_DisplayNodes(lID)).Path)
                            If m_Nodes(m_DisplayNodes(lID)).State = ens_NotExpanded Then
                                pvAddFolders m_DisplayNodes(lID)

                                '-- Set The Expanded Flag
                                m_Nodes(m_DisplayNodes(lID)).State = ens_ExpandedOnce
                                UserControl.MousePointer = vbDefault
                                m_Nodes(m_DisplayNodes(lID)).Expanded = True

                                '-- Make the DisplayNodes
                                pvMakeDisplayNodes

                                If Not pbIsDimmed(hTable(m_Nodes(m_DisplayNodes(lID)).Children)) Then
                                    m_Nodes(m_DisplayNodes(lID)).HasChildren = False
                                Else
                                    m_Nodes(m_DisplayNodes(lID)).ItemData = "(" & UBound(m_Nodes(m_DisplayNodes(lID)).Children) + 1 & ")"
                                End If

                            End If

                            RaiseEvent LocatingComplete


                        Else
                            m_Nodes(m_DisplayNodes(lID)).Expanded = Not m_Nodes(m_DisplayNodes(lID)).Expanded
                        End If
                    Else


                        If m_lType = edt_Custom Then

                            If m_Nodes(m_DisplayNodes(lID)).HasChildren Then
                                If Not m_Nodes(m_DisplayNodes(lID)).State = ens_ExpandedOnce Then
                                    bExpanded = Not m_Nodes(m_DisplayNodes(lID)).Expanded
                                    sCap = m_Nodes(m_DisplayNodes(lID)).Caption
                                    RaiseEvent BeforeCustomNodeExpand(lID, sCap, bExpanded)
                                    m_Nodes(m_DisplayNodes(lID)).State = ens_ExpandedOnce
                                End If
                            End If
                        End If

                        m_Nodes(m_DisplayNodes(lID)).Expanded = Not m_Nodes(m_DisplayNodes(lID)).Expanded
                        '-- Make the DisplayNodes
                    End If
                End If
            End If
            pvRedraw


        End If

    ElseIf Button = vbRightButton Then
        If lID > 0 Then RaiseEvent RightClick(m_Nodes(m_DisplayNodes(lID)).Path)
    End If

    '-- Make the DisplayNodes

    pvMakeDisplayNodes
    pvRedraw
    UserControl.MousePointer = vbDefault
    'Debug.Print m_Nodes(m_DisplayNodes(lID)).Path


End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lLasty = Y
    llastx = X
    Dim lIID As Long
    Dim bChevron As Boolean

    If Me.HoverSelection Then
        pvHittest X, Y, lIID, bChevron
        m_HoverId = m_DisplayNodes(lIID)
        DoEvents
        pvRedraw
    End If

End Sub


Private Sub UserControl_Resize()
    On Error Resume Next

    pvMakeScrollScrollbarVisible
    pvRedraw

    On Error GoTo 0
End Sub

Private Sub UserControl_Terminate()

'-- Ensure all GDI objects are freed:

    pvClearUp
    '-- Clear up the picture:
    Set m_pic = Nothing

End Sub

Private Sub VScroll1_Change()

    DoEvents
    pvRedraw

End Sub

Private Sub VScroll1_Scroll()
    DoEvents
    pvRedraw
End Sub
'//---------------------------------------------------------------------------------------
' Procedure : pvAddDrives
' Type      : Sub
' DateTime  : 06/10/2004 17:34
' Author    : Gary Noble
' Purpose   : Adds The Drives To The Node Array
' Returns   :
' Notes     : Taken From CarlesPV's Directory Tree Sample, All Credit To Him!
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvAddDrives()

    Dim uSHFI          As SHFILEINFO
    Dim lPIDL          As Long
    Dim sBuffer        As String * MAX_PATH
    Dim lDrivesBitMask As Long
    Dim lMaxPwr        As Long
    Dim lPwr           As Long
    Dim hNodeRoot      As Long
    Dim hNode          As Long
    Dim sText          As String
    Dim lRet           As Long





    Call SHGetSpecialFolderLocation(0, CSIDL_PERSONAL, lPIDL)
    Call SHGetPathFromIDListA(lPIDL, sBuffer)
    Call SHGetFileInfo(ByVal lPIDL, 0, uSHFI, Len(uSHFI), SHGFI_PIDL + SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX)
    Call CoTaskMemFree(lPIDL)
    hNode = pvTVAdd(0, , uSHFI.szDisplayName, lMaxPwr + 1, uSHFI.iIcon, uSHFI.iIcon, bForcePlusButton:=True)
    m_Nodes(hNode).Path = pvStripNulls(sBuffer) & "\"
    m_Nodes(hNode).Icon = uSHFI.iIcon
    '    m_Nodes(hNodeRoot).ItemData = "(" & UBound(m_Nodes(hNodeRoot).Children) + 1 & ")"

    '-- Add root node ('My Computer') (PIDL)

    Call SHGetSpecialFolderLocation(0, CSIDL_DRIVES, lPIDL)
    Call SHGetFileInfo(ByVal lPIDL, 0, uSHFI, Len(uSHFI), SHGFI_PIDL + SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX)
    Call CoTaskMemFree(lPIDL)
    hNodeRoot = pvTVAdd(, , uSHFI.szDisplayName, 0, uSHFI.iIcon, uSHFI.iIcon, True, uSHFI.szDisplayName)
    m_Nodes(hNodeRoot).Expanded = True
    m_Nodes(hNodeRoot).Icon = uSHFI.iIcon
    m_Nodes(hNodeRoot).HasChildren = True
    m_Nodes(hNodeRoot).State = ens_ExpandedOnce
    '-- Add drives

    lDrivesBitMask = GetLogicalDrives()

    If (lDrivesBitMask) Then

        lMaxPwr = Int(Log(lDrivesBitMask) / Log(2))


        For lPwr = 0 To lMaxPwr

            If (2 ^ lPwr And lDrivesBitMask) Then

                sText = Chr$(65 + lPwr) & ":\"

                lRet = SHGetFileInfo(sText, 0, uSHFI, Len(uSHFI), SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX)
                hNode = pvTVAdd(hNodeRoot, , uSHFI.szDisplayName, lPwr, uSHFI.iIcon, uSHFI.iIcon, bForcePlusButton:=True, sPath:=sText)
                m_Nodes(hNode).HasChildren = True
                m_Nodes(hNode).Icon = uSHFI.iIcon
            End If
        Next lPwr
    End If




End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvAddFolders
' Type      : Sub
' DateTime  : 06/10/2004 17:35
' Author    : Gary Noble
' Purpose   : Attaches The Folders To The Directory
' Returns   :
' Notes     : Taken From CarlesPV's Directory Tree Sample, All Credit To Him!
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvAddFolders(ByVal hNode As Long)

    Dim uSHFI       As SHFILEINFO
    Dim uWFD        As WIN32_FIND_DATA
    Dim lParam      As Long
    Dim lNextIdx    As Long
    Dim sPath       As String
    Dim sFolderName As String
    Dim lFolders    As Long
    Dim hSearch     As Long
    Dim hNext       As Long
    Dim lAttr       As Long
    Dim lRet        As Long

    ' If (pvTVGetRoot() <> hNode) Then
    sPath = m_Nodes(hNode).Path
    '    MsgBox sPath


    '-- Start searching
    hNext = 1
    hSearch = FindFirstFile(sPath & "*.", uWFD)

    If (hSearch <> INVALID_HANDLE_VALUE) Then

        Do While hNext

            '-- Get file [folder] name
            sFolderName = pvStripNulls(uWFD.cFileName)

            If (sFolderName <> "." And sFolderName <> "..") Then

                '-- Only standard folders
                lAttr = GetFileAttributes(sPath & sFolderName)
                If (lAttr = FILE_ATTRIBUTE_DIRECTORY Or _
                        lAttr = FILE_ATTRIBUTE_RODIRECTORY) Then

                    '-- Get info (name and image list index)
                    lRet = SHGetFileInfo(sPath & sFolderName & "\", 0, uSHFI, Len(uSHFI), SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX)

                    '-- Add node
                    lRet = pvTVAdd(hNode, , uSHFI.szDisplayName, lNextIdx, uSHFI.iIcon, uSHFI.iIcon + -(uSHFI.iIcon = SI_FOLDER_CLOSED), -pvHasSubFolders(sPath & sFolderName & "\"), sPath & sFolderName & "\")
                    m_Nodes(lRet).HasChildren = pvHasSubFolders(sPath & sFolderName & "\")
                    m_Nodes(lRet).Icon = uSHFI.iIcon
                    '-- Count folders (-> Sort children ?)
                    lFolders = lFolders + 1
                End If
            End If
            hNext = FindNextFile(hSearch, uWFD)
        Loop
        hNext = FindClose(hSearch)

        '-- Sort added folders and ensure ScrollbarVisible parent
        'Call pvTVSortChildren(hNode)
        'Call pvTVEnsureScrollbarVisible(hNode)
        'End If

        '-- Hide 'plus button' ?
        If (lFolders = 0) Then
            '            Call pvTVSetcChildren(hNode, 0)
        End If
    End If
    pvMakeDisplayNodes
End Sub


'//---------------------------------------------------------------------------------------
' Procedure : pvHasSubFolders
' Type      : Function
' DateTime  : 06/10/2004 17:35
' Author    : Gary Noble
' Purpose   : Sets The has Folders Flag
' Returns   : Boolean
' Notes     : Taken From CarlesPV's Directory Tree Sample, All Credit To Him!
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Function pvHasSubFolders(ByVal sPath As String) As Boolean

    Dim uWFD        As WIN32_FIND_DATA
    Dim sFolderName As String
    Dim hSearch     As Long
    Dim hNext       As Long
    Dim lAttr       As Long

    '-- Start searching
    hNext = 1
    hSearch = FindFirstFile(sPath & "*.", uWFD)

    If (hSearch <> INVALID_HANDLE_VALUE) Then

        Do While hNext

            sFolderName = pvStripNulls(uWFD.cFileName)
            If (sFolderName <> "." And sFolderName <> "..") Then

                lAttr = GetFileAttributes(sPath & sFolderName)
                If (lAttr = FILE_ATTRIBUTE_DIRECTORY Or _
                        lAttr = FILE_ATTRIBUTE_RODIRECTORY) Then

                    '-- Found one: enough
                    pvHasSubFolders = True
                    Exit Do
                End If
            End If
            hNext = FindNextFile(hSearch, uWFD)
        Loop
        hNext = FindClose(hSearch)
    End If
End Function


'//---------------------------------------------------------------------------------------
' Procedure : pvTVAdd
' Type      : Function
' DateTime  : 06/10/2004 17:36
' Author    : Gary Noble
' Purpose   : Global Calling Point For Adding a node To The array
' Returns   : Long
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Function pvTVAdd( _
                 Optional ByVal hParent As Long = 0, _
                 Optional ByVal hInsertAfter As Long = -1, _
                 Optional ByVal sText As String = vbNullString, _
                 Optional ByVal lParam As Long = -1, _
                 Optional ByVal lImage As Long = -1, _
                 Optional ByVal lSelectedImage As Long = -1, _
                 Optional ByVal bForcePlusButton As Boolean = False _
                 , Optional sPath As String) As Long

    pvTVAdd = plAddNode(RTrim(pvStripNulls(sText)), RTrim(pvStripNulls(sPath)), 0, hParent)
    m_Nodes(pvTVAdd).HasChildren = bForcePlusButton


End Function

Private Function pvGetSystemImageList(ByVal uSize As Long) As Long
    Dim uSHFI As SHFILEINFO
    m_lptrVb6ImageList = SHGetFileInfo("C:\", 0, uSHFI, Len(uSHFI), SHGFI_SYSICONINDEX Or uSize)
End Function


'//---------------------------------------------------------------------------------------
' Procedure : pvStripNulls
' Type      : Function
' DateTime  : 06/10/2004 17:36
' Author    : Gary Noble
' Purpose   : Strips nulls From A String
' Returns   : String
' Notes     : Taken From CarlesPV's Directory Tree Sample, All Credit To Him!
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Function pvStripNulls(ByVal sString As String) As String

    Dim lPos As Long

    lPos = InStr(sString, vbNullChar)

    If (lPos = 1) Then
        pvStripNulls = vbNullString
    ElseIf (lPos > 1) Then
        pvStripNulls = Left$(sString, lPos - 1)
        Exit Function
    End If

    pvStripNulls = sString
End Function


'//---------------------------------------------------------------------------------------
' Procedure : pvImageListDrawIcon
' Type      : Sub
' DateTime  : 06/10/2004 17:36
' Author    : Gary Noble
' Purpose   : Draws A Image To The DC From A Image List
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvImageListDrawIcon(ByVal ptrVb6ImageList As Long, _
                              ByVal lngHdc As Long, _
                              ByVal hIml As Long, _
                              ByVal iIconIndex As Long, _
                              ByVal lX As Long, _
                              ByVal lY As Long, _
                              Optional ByVal bSelected As Boolean = False, _
                              Optional ByVal bBlend25 As Boolean = False, _
                              Optional ByVal IsHeaderIcon As Boolean = False)


    Dim o          As Object
    Dim lFlags     As Long
    Dim lR         As Long
    Dim icoInfo    As ICONINFO
    Dim newICOinfo As ICONINFO
    Dim icoBMPinfo As BITMAP

    If Not Me.Enabled Then
        pvImageListDrawIconDisabled ptrVb6ImageList, lngHdc, hIml, iIconIndex, lX, lY, m_lIconHeight, True
        Exit Sub
    End If
    lFlags = ILD_TRANSPARENT
    If bSelected Then
        lFlags = lFlags Or ILD_SELECTED
    End If
    If bBlend25 Then
        lFlags = lFlags Or ILD_BLEND25
    End If
    If ptrVb6ImageList <> 0 Then
        On Error Resume Next
        lR = ImageList_Draw(ptrVb6ImageList, iIconIndex, lngHdc, lX, lY, lFlags)
        If lR = 0 Then
            'Debug.Print "Failed to draw Image: " & iIconIndex & " onto hDC " & hDC, "pvImageListDrawIcon"
        End If
    End If

End Sub


'//---------------------------------------------------------------------------------------
' Procedure : pvImageListDrawIconDisabled
' Type      : Sub
' DateTime  : 06/10/2004 17:37
' Author    : Gary Noble
' Purpose   : Draws A Disabled image List Icon
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvImageListDrawIconDisabled(ByVal ptrVb6ImageList As Long, _
                                      ByVal lngHdc As Long, _
                                      ByVal hIml As Long, _
                                      ByVal iIconIndex As Long, _
                                      ByVal lX As Long, _
                                      ByVal lY As Long, _
                                      ByVal lSize As Long, _
                                      Optional ByVal asShadow As Boolean)

    Dim o     As Object
    Dim hBr   As Long
    Dim hIcon As Long

    'Dim lR    As Long
    hIcon = 0
    If ptrVb6ImageList <> 0 Then
        On Error Resume Next
        Set o = ObjectFromPtr(ptrVb6ImageList)
        If Not (o Is Nothing) Then
            hIcon = o.ListImages(iIconIndex + 1).ExtractIcon()
        End If
        On Error GoTo 0
    Else
        hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
    End If
    If hIcon <> 0 Then
        If asShadow Then
            hBr = GetSysColorBrush(vb3DShadow And &H1F)
            If lngHdc = hDC Then
                DrawState lngHdc, hBr, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_MONO
            Else
                DrawState lngHdc, hBr, 0, hIcon, 0, lX, lY + 4, 16, 16, DST_ICON Or DSS_MONO
            End If
            DeleteObject hBr
        Else
            DrawState lngHdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED
        End If
        DestroyIcon hIcon
    End If

End Sub


Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Attribute Font.VB_Description = "Returns a Font object."
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_HoverSelection = m_def_HoverSelection
    Set m_BackGroundPicture = LoadPicture("")


End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If UserControl.Ambient.UserMode Then
        Set m_cScrollBar = New IAPP_ScrollBars
        m_cScrollBar.Create UserControl.hwnd
        m_cScrollBar.Orientation = efsoVertical
        m_cScrollBar.Visible(efsVertical) = False
        m_cScrollBar.Visible(efsHorizontal) = False
    End If

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_HoverSelection = PropBag.ReadProperty("HoverSelection", m_def_HoverSelection)
    Set m_BackGroundPicture = PropBag.ReadProperty("BackGroundPicture", Nothing)

    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("HoverSelection", m_HoverSelection, m_def_HoverSelection)
    Call PropBag.WriteProperty("BackGroundPicture", m_BackGroundPicture, Nothing)

    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub


'//---------------------------------------------------------------------------------------
' Procedure : pvHittest
' Type      : Sub
' DateTime  : 06/10/2004 17:37
' Author    : Gary Noble
' Purpose   : Did We Hit A Node
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvHittest(X As Single, Y As Single, rID As Long, bChevron As Boolean)

    Dim i As Long
    Dim oRect As RECT


    For i = m_cScrollBar.Value(efsVertical) + 1 To m_lLastIDDrawn
        If PtInRect(m_Nodes(m_DisplayNodes(i)).oRect, X, Y) Then
            If X < m_Nodes(m_DisplayNodes(i)).oRect.Left + 12 Then bChevron = True
            rID = i
            'm_lSelectionID = i
            Exit For
        End If
    Next


End Sub

Public Property Get HoverSelection() As Boolean
    HoverSelection = m_HoverSelection
End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)
    m_HoverSelection = New_HoverSelection
    PropertyChanged "HoverSelection"
End Property

'//---------------------------------------------------------------------------------------
' Procedure : pvDisplayStartupPath
' Type      : Function
' DateTime  : 06/10/2004 17:38
' Author    : Gary Noble
' Purpose   : Loops Through The Nodes Adding The Folders If Neccessary
'             The Finds The Selected Path and Selects It.
' Returns   : Boolean
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Function pvDisplayStartupPath(sPath As String, Optional lSelID As Long) As Boolean

    Dim i As Long
    Dim iFind As Long
    Dim vPaths As Variant
    Dim sPathNew As String
    Dim iFound As Long


    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"

    vPaths = Split(sPath, "\")

    Dim ll As Long


    For i = 0 To UBound(vPaths) - 1
        sPathNew = sPathNew & vPaths(i) & "\"
        For iFind = 1 To UBound(m_DisplayNodes)
            If LCase(m_Nodes(m_DisplayNodes(iFind)).Path) = LCase(sPathNew) Then
                ll = ll + 1

                If m_Nodes(m_DisplayNodes(iFind)).State = ens_ExpandedOnce Then
                    m_Nodes(m_DisplayNodes(iFind)).Expanded = True
                    m_Nodes(m_DisplayNodes(iFind)).State = ens_ExpandedOnce
                    If Not pbIsDimmed(hTable(m_Nodes(m_DisplayNodes(iFind)).Children)) Then
                        m_Nodes(m_DisplayNodes(iFind)).HasChildren = False
                    Else
                        m_Nodes(m_DisplayNodes(iFind)).ItemData = "(" & UBound(m_Nodes(m_DisplayNodes(iFind)).Children) + 1 & ")"
                    End If
                    m_lSelectedID = m_DisplayNodes(iFind)
                    pvMakeDisplayNodes
                    Exit For

                Else
                    ll = ll + 1
                    m_Nodes(m_DisplayNodes(iFind)).Expanded = True
                    m_Nodes(m_DisplayNodes(iFind)).State = ens_ExpandedOnce
                    pvAddFolders m_DisplayNodes(iFind)
                    If Not pbIsDimmed(hTable(m_Nodes(m_DisplayNodes(iFind)).Children)) Then
                        m_Nodes(m_DisplayNodes(iFind)).HasChildren = False
                    Else
                        m_Nodes(m_DisplayNodes(iFind)).ItemData = "(" & UBound(m_Nodes(m_DisplayNodes(iFind)).Children) + 1 & ")"
                    End If
                    m_lSelectedID = m_DisplayNodes(iFind)
                    pvMakeDisplayNodes

                    Exit For


                End If
            End If
        Next
    Next

YeahOK:
    On Error Resume Next
    iFound = 0

    m_cScrollBar.Value(efsVertical) = 0
    SendMessage hwnd, WM_SETREDRAW, 0, 0

    On Error Resume Next
    m_bDrawingSelectedNode = False

    While Not m_bDrawingSelectedNode
        '  GoTo CleanExit

        If m_bDrawingSelectedNode Then GoTo CleanExit
        m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) + 1    ' + (llastdrawn - lfirstdrawn) - 3
        If m_cScrollBar.Value(efsVertical) >= m_cScrollBar.Max(efsVertical) Then GoTo CleanExit    '
    Wend

    pvRedraw

CleanExit:

    SendMessage hwnd, WM_SETREDRAW, 1, 0

    m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical)


    RaiseEvent SelectionChanged(m_Nodes(m_DisplayNodes(iFind)).Caption, m_Nodes(m_DisplayNodes(iFind)).Path)
    pvRedraw

End Function


'//---------------------------------------------------------------------------------------
' Procedure : SelectPath
' Type      : Sub
' DateTime  : 06/10/2004 17:39
' Author    : Gary Noble
' Purpose   : Finds A Path in the Tree
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Public Sub SelectPath(ByVal sPath As String)

    m_bDrawingSelectedNode = False
    If Len(Trim(sPath)) > 0 Then pvDisplayStartupPath sPath

End Sub


'-----------------------------------------------------------
'-- Drawing Routines


'//---------------------------------------------------------------------------------------
' Procedure : pvDrawOpenCloseGlyph
' Type      : Sub
' DateTime  : 06/10/2004 17:39
' Author    : Gary Noble
' Purpose   : Draws The Expand/Collapse Button
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvDrawOpenCloseGlyph(ByVal lngHwnd As Long, _
                              ByVal lHDC As Long, _
                              tTR As RECT, _
                              ByVal bCollapsed As Boolean)


    Dim tGR     As RECT
    Dim bDone   As Boolean
    Dim hTheme  As Long
    Dim hBr     As Long
    Dim hPen    As Long
    Dim hPenOld As Long
    Dim tJ      As POINTAPI

    LSet tGR = tTR

    With tGR
        .Left = .Left + 2
        .Right = .Left + 12
        .Top = .Top + 2    ' ((.bottom - .top) \ 2)
        .Bottom = .Top + 12
    End With    'tGR


    If IsXp Then
        hTheme = OpenThemeData(lngHwnd, StrPtr("TREEVIEW"))
        If Not (hTheme = 0) Then
            DrawThemeBackground hTheme, lHDC, 2, IIf(bCollapsed, 1, 2), tGR, tGR
            CloseThemeData hTheme
            bDone = True
        End If
    End If

    If Not (bDone) Then
        '-- Draw button border
        ' hBr = GetSysColorBrush(vbButtonFace And &H1F&)
        ' FillRect lHDC, tGR, hBr
        ' DeleteObject hBr
        hPen = CreatePen(PS_SOLID, 1, plTranslateColor(vbWindowText))
        hPenOld = SelectObject(lHDC, hPen)
        '

        pvDrawBorderRectangle lHDC, vbBlack, tGR.Left + 1, tGR.Top + 1, 9, 9, False

        SelectObject lHDC, hPenOld
        DeleteObject hPen
        hPen = CreatePen(PS_SOLID, 1, plTranslateColor(vbWindowText))
        hPenOld = SelectObject(lHDC, hPen)
        '
        With tGR
            MoveToEx lHDC, .Left + 3, .Top + 5, tJ
            LineTo lHDC, .Left + 8, .Top + 5
        End With    'tGR
        If bCollapsed Then
            MoveToEx lHDC, tGR.Left + 5, tGR.Top + 3, tJ
            LineTo lHDC, tGR.Left + 5, tGR.Top + 8
        End If
        SelectObject lHDC, hPenOld
        DeleteObject hPen
    End If

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvGradientFillRect
' Type      : Sub
' DateTime  : 06/10/2004 17:40
' Author    : Gary Noble
' Purpose   : Fills A Rect With A Two Tone Gradient
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvGradientFillRect(ByVal lHDC As Long, _
                             tR As RECT, _
                             ByVal oStartColor As OLE_COLOR, _
                             ByVal oEndColor As OLE_COLOR, _
                             ByVal eDir As GradientFillRectType)

    Dim tTV(0 To 1) As TRIVERTEX
    Dim tGR         As GRADIENT_RECT
    Dim hBrush      As Long
    Dim lStartColor As Long
    Dim lEndColor   As Long

    'Dim lR As Long
    '-- Use GradientFill:
    If HasGradientAndTransparency Then
        lStartColor = plTranslateColor(oStartColor)
        lEndColor = plTranslateColor(oEndColor)
        pvsetTriVertexColor tTV(0), lStartColor
        tTV(0).X = tR.Left
        tTV(0).Y = tR.Top
        pvsetTriVertexColor tTV(1), lEndColor
        tTV(1).X = tR.Right
        tTV(1).Y = tR.Bottom
        tGR.UpperLeft = 0
        tGR.LowerRight = 1
        GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
    Else
        '-- Fill with solid brush:
        hBrush = CreateSolidBrush(plTranslateColor(oEndColor))
        FillRect lHDC, tR, hBrush
        DeleteObject hBrush
    End If

End Sub

Private Property Get HasGradientAndTransparency()


    HasGradientAndTransparency = m_bHasGradientAndTransparency

End Property

Private Property Get Is2000OrAbove() As Boolean


    Is2000OrAbove = m_bIs2000OrAbove

End Property

Private Property Get IsNt() As Boolean


    IsNt = m_bIsNt

End Property

Private Property Get IsXp() As Boolean

    IsXp = m_bIsXp

End Property

'-- Gradient Call
Private Sub pvsetTriVertexColor(tTV As TRIVERTEX, _
                              ByVal lColor As Long)

    Dim lRed   As Long
    Dim lGreen As Long
    Dim lBlue  As Long

    lRed = (lColor And &HFF&) * &H100&
    lGreen = (lColor And &HFF00&)
    lBlue = (lColor And &HFF0000) \ &H100&
    With tTV
        pvsetTriVertexColorComponent .Red, lRed
        pvsetTriVertexColorComponent .Green, lGreen
        pvsetTriVertexColorComponent .Blue, lBlue
    End With    'tTV

End Sub

Private Sub pvsetTriVertexColorComponent(ByRef iColor As Integer, _
                                       ByVal lComponent As Long)

    If (lComponent And &H8000&) = &H8000& Then
        iColor = (lComponent And &H7F00&)
        iColor = iColor Or &H8000
    Else
        iColor = lComponent
    End If

End Sub

Private Sub pvDrawBackground(ByVal lngHdc As Long, _
                              ByVal colorStart As Long, _
                              ByVal colorEnd As Long, _
                              ByVal lngLeft As Long, _
                              ByVal lngTop As Long, _
                              ByVal lngWidth As Long, _
                              ByVal lngHeight As Long, _
                              Optional ByVal horizontal As Boolean = False)

    Dim tR As RECT

    With tR
        .Left = lngLeft
        .Top = lngTop
        .Right = lngLeft + lngWidth
        .Bottom = lngTop + lngHeight
        '-- gradient fill vertical:
    End With    'tR
    pvGradientFillRect lngHdc, tR, colorStart, colorEnd, IIf(horizontal, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)

End Sub


'//---------------------------------------------------------------------------------------
' Procedure : pvDrawBorderRectangle
' Type      : Sub
' DateTime  : 06/10/2004 17:41
' Author    : Gary Noble
' Purpose   : Draws A Rectangle
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvDrawBorderRectangle(ByVal lngHdc As Long, _
                                   ByVal lColor As Long, _
                                   ByVal lngLeft As Long, _
                                   ByVal lngTop As Long, _
                                   ByVal lngWidth As Long, _
                                   ByVal lngHeight As Long, _
                                   ByVal bInset As Boolean)


    Dim tJ      As POINTAPI
    Dim hPen    As Long
    Dim hPenOld As Long

    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lngHdc, hPen)
    MoveToEx lngHdc, lngLeft, lngTop + lngHeight - 1, tJ
    LineTo lngHdc, lngLeft, lngTop
    LineTo lngHdc, lngLeft + lngWidth - 1, lngTop
    LineTo lngHdc, lngLeft + lngWidth - 1, lngTop + lngHeight - 1
    LineTo lngHdc, lngLeft, lngTop + lngHeight - 1
    SelectObject lngHdc, hPenOld
    DeleteObject hPen

End Sub


'//---------------------------------------------------------------------------------------
' Procedure : pvDrawText
' Type      : Sub
' DateTime  : 06/10/2004 17:41
' Author    : Gary Noble
' Purpose   : Draws Text To The Give DC
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvDrawText(ByVal lngHdc As Long, _
                        ByVal sCaption As String, _
                        ByVal lTextX As Long, _
                        ByVal lTextY As Long, _
                        ByVal lTextX1 As Long, _
                        ByVal lTextY1 As Long, _
                        ByVal bEnabled As Boolean, _
                        ByVal color As Long, _
                        ByVal bCentreHorizontal As Boolean, _
                        Optional RightAlign As Boolean = False)


    Dim rcText As RECT

    SetTextColor lngHdc, color
    'Dim lFlags As Long
    If Not bEnabled Then
        SetTextColor lngHdc, GetSysColor(vbGrayText And &H1F&)
    End If
    With rcText
        .Left = lTextX
        .Top = lTextY
        .Right = lTextX1
        .Bottom = lTextY1
    End With
    If m_bIsNt Then
        DrawTextW lngHdc, StrPtr(sCaption), -1, rcText, IIf(RightAlign, DT_RIGHT, DT_LEFT) Or DT_END_ELLIPSIS
    Else
        DrawTextA lngHdc, sCaption, -1, rcText, IIf(RightAlign, DT_RIGHT, DT_LEFT) Or DT_END_ELLIPSIS
    End If
    If Not bEnabled Then
        SetTextColor lngHdc, plTranslateColor(vbWindowText)
    End If

End Sub


'//---------------------------------------------------------------------------------------
' Procedure : pvSetCursor
' Type      : Sub
' DateTime  : 06/10/2004 17:41
' Author    : Gary Noble
' Purpose   : Sets The Cursor To The windows Default Or Hand Cursor
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvSetCursor(ByVal bHand As Boolean)

'-- Desc: Get the "Real" Hand Cursor

    If bHand Then
        SetCursor LoadCursor(0, IDC_HAND)
        m_bHandCursor = True
    Else
        SetCursor LoadCursor(0, IDC_ARROW)
        m_bHandCursor = False
    End If

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvVerInitialise
' Type      : Sub
' DateTime  : 06/10/2004 17:41
' Author    : Gary Noble
' Purpose   : Environmental Settings
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvVerInitialise()

    Dim tOSV As OSVERSIONINFO

    tOSV.dwVersionInfoSize = Len(tOSV)
    GetVersionEx tOSV
    m_bIsNt = ((tOSV.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    If tOSV.dwMajorVersion > 5 Then
        m_bHasGradientAndTransparency = True
        m_bIsXp = True
        m_bIs2000OrAbove = True
    ElseIf (tOSV.dwMajorVersion = 5) Then
        m_bHasGradientAndTransparency = True
        m_bIs2000OrAbove = True
        If tOSV.dwMinorVersion >= 1 Then
            m_bIsXp = True
        End If
    ElseIf (tOSV.dwMajorVersion = 4) Then    '-- NT4 or 9x/ME/SE
        If tOSV.dwMinorVersion >= 10 Then
            m_bHasGradientAndTransparency = True
        End If
    Else    '-- Too old
    End If

End Sub


'//---------------------------------------------------------------------------------------
' Procedure : BlendColor
' Type      : Property
' DateTime  : 06/10/2004 17:42
' Author    : Gary Noble
' Purpose   : Blends 2 Colors Together
' Returns   : Long
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, _
                               ByVal oColorTo As OLE_COLOR) As Long


    Dim lCFrom As Long
    Dim lCTo   As Long
    Dim lCRetR As Long
    Dim lCRetG As Long
    Dim lCRetB As Long

    lCFrom = plTranslateColor(oColorFrom)
    lCTo = plTranslateColor(oColorTo)
    lCRetR = (lCFrom And &HFF) + ((lCTo And &HFF) - (lCFrom And &HFF)) \ 2
    If lCRetR > 255 Then
        lCRetR = 255
    ElseIf (lCRetR < 0) Then
        lCRetR = 0
    End If
    lCRetG = ((lCFrom \ &H100) And &HFF&) + (((lCTo \ &H100) And &HFF&) - ((lCFrom \ &H100) And &HFF&)) \ 2
    If lCRetG > 255 Then
        lCRetG = 255
    ElseIf (lCRetG < 0) Then
        lCRetG = 0
    End If
    lCRetB = ((lCFrom \ &H10000) And &HFF&) + (((lCTo \ &H10000) And &HFF&) - ((lCFrom \ &H10000) And &HFF&)) \ 2
    If lCRetB > 255 Then
        lCRetB = 255
    ElseIf (lCRetB < 0) Then
        lCRetB = 0
    End If
    BlendColor = RGB(lCRetR, lCRetG, lCRetB)

End Property


Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object

    Dim oTemp As Object

    '-- Turn the pointer into an illegal, uncounted interface
    CopyMemory oTemp, lPtr, 4
    '-- Do NOT hit the End button here! You will crash!
    '-- Assign to legal reference
    Set ObjectFromPtr = oTemp
    '-- Still do NOT hit the End button here! You will still crash!
    '-- Destroy the illegal reference
    CopyMemory oTemp, 0&, 4
    '-- OK, hit the End button if you must--you'll probably still crash,
    '-- but it will be because of the subclass, not the uncounted reference

End Property

Private Function plTranslateColor(ByVal oClr As OLE_COLOR, _
                               Optional hPal As Long = 0) As Long

'-- Convert Automation color to Windows color

    If OleTranslateColor(oClr, hPal, plTranslateColor) Then
        plTranslateColor = CLR_INVALID
    End If

End Function

Public Property Get BackGroundPicture() As Picture
    Set BackGroundPicture = m_BackGroundPicture
End Property

Public Property Set BackGroundPicture(ByVal New_BackGroundPicture As Picture)
    Set m_BackGroundPicture = New_BackGroundPicture
    PropertyChanged "BackGroundPicture"
    pvRedraw
    Refresh
End Property

'-- Tiler


Private Property Get BitmapHeight() As Long


    BitmapHeight = m_lBitmapH

End Property

Private Property Get BitmapWidth() As Long


    BitmapWidth = m_lBitmapW

End Property





Private Function pbEnsurePicture() As Boolean

    On Error Resume Next
    pbEnsurePicture = True
    If (m_pic Is Nothing) Then
        Set m_pic = New StdPicture
        If Err.Number <> 0 Then
            'pErr 3, "Unable to allocate memory for picture object."
            pbEnsurePicture = False
        Else
        End If
    End If
    On Error GoTo 0

End Function


'//---------------------------------------------------------------------------------------
' Procedure : pbGetBitmapIntoDC
' Type      : Function
' DateTime  : 06/10/2004 17:42
' Author    : Gary Noble
' Purpose   : Puts A Bitmap In To A DC Ready For Calling To Draw
' Returns   : Boolean
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Function pbGetBitmapIntoDC() As Boolean

    Dim tB           As BITMAP
    Dim lHDC         As Long
    Dim lHwnd        As Long
    Dim lHDCTemp     As Long
    Dim lHBmpTempOld As Long

    '-- Make a DC to hold the picture bitmap which we can blt from:
    lHwnd = GetDesktopWindow()
    lHDC = GetDC(lHwnd)
    m_lHdc = CreateCompatibleDC(lHDC)
    lHDCTemp = CreateCompatibleDC(lHDC)
    If m_lHdc <> 0 Then
        '-- Get size of bitmap:
        GetObjectAPI m_pic.Handle, LenB(tB), tB
        m_lBitmapW = tB.bmWidth
        m_lBitmapH = tB.bmHeight
        lHBmpTempOld = SelectObject(lHDCTemp, m_pic.Handle)
        m_lHBmp = CreateCompatibleBitmap(lHDC, m_lBitmapW, m_lBitmapH)
        m_lHBmpOld = SelectObject(m_lHdc, m_lHBmp)
        BitBlt m_lHdc, 0, 0, m_lBitmapW, m_lBitmapH, lHDCTemp, 0, 0, vbSrcCopy
        SelectObject lHDCTemp, lHBmpTempOld
        DeleteDC lHDCTemp
        If m_lHBmpOld <> 0 Then
            pbGetBitmapIntoDC = True
            If LenB(m_sFileName) = 0 Then
                m_sFileName = "PICTURE"
            End If
        Else
            pvClearUp
            'pErr 2, "Unable to select bitmap into DC"
        End If
    Else
        'pErr 1, "Unable to create compatible DC"
    End If
    ReleaseDC lHwnd, lHDC

End Function


'//---------------------------------------------------------------------------------------
' Procedure : pvClearUp
' Type      : Sub
' DateTime  : 06/10/2004 17:43
' Author    : Gary Noble
' Purpose   : DC/GDI Cleanup Routine
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvClearUp()

'-- Clear reference to the filename:

    m_sFileName = ""
    '-- If we have a DC, then clear up:
    If m_lHdc <> 0 Then
        '-- Select the bitmap out of DC:
        If m_lHBmpOld <> 0 Then
            SelectObject m_lHdc, m_lHBmpOld
            '-- The original bitmap does not have to deleted because it is owned by m_pic
        End If
        If m_lHBmp <> 0 Then
            DeleteObject m_lHBmp
        End If
        '-- Remove the DC:
        DeleteDC m_lHdc
    End If

End Sub



Private Property Get pPicture() As StdPicture

    Set pPicture = m_pic

End Property

Private Sub pvSetPicture(oPic As StdPicture)

'-- Load a picture from a StdPicture object:

    pvClearUp
    If Not oPic Is Nothing Then
        If pbEnsurePicture() Then
            Set m_pic = oPic
            If Err.Number = 0 Then
                pbGetBitmapIntoDC
            End If
        End If
    End If

End Sub




'//---------------------------------------------------------------------------------------
' Procedure : pvTileArea
' Type      : Sub
' DateTime  : 06/10/2004 17:43
' Author    : Gary Noble
' Purpose   : Tiles A Picture To the DC
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvTileArea(ByRef lngHdc As Long, _
                    ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal lngWidth As Long, _
                    ByVal lngHeight As Long)


    Dim lSrcX           As Long
    Dim lSrcY           As Long
    Dim lSrcStartX      As Long
    Dim lSrcStartY      As Long
    Dim lSrcStartWidth  As Long
    Dim lSrcStartHeight As Long
    Dim lDstX           As Long
    Dim lDstY           As Long
    Dim lDstWidth       As Long
    Dim lDstHeight      As Long

    lSrcStartX = ((X + m_lXOriginOffset) Mod m_lBitmapW)
    lSrcStartY = ((Y + m_lYOriginOffset) Mod m_lBitmapH)
    lSrcStartWidth = (m_lBitmapW - lSrcStartX)
    lSrcStartHeight = (m_lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    lDstY = Y
    lDstHeight = lSrcStartHeight
    Do While lDstY < (Y + lngHeight)
        If (lDstY + lDstHeight) > (Y + lngHeight) Then
            lDstHeight = Y + lngHeight - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = X
        lSrcX = lSrcStartX
        Do While lDstX < (X + lngWidth)
            If (lDstX + lDstWidth) > (X + lngWidth) Then
                lDstWidth = X + lngWidth - lDstX
                If lDstWidth = 0 Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt lngHdc, lDstX, lDstY, lDstWidth, lDstHeight, m_lHdc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = m_lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = m_lBitmapH
    Loop

End Sub

Private Property Get XOriginOffset() As Long

    XOriginOffset = m_lXOriginOffset

End Property

Private Property Let XOriginOffset(ByVal lPixels As Long)

    m_lXOriginOffset = lPixels

End Property

Private Property Get YOriginOffset() As Long

    YOriginOffset = m_lYOriginOffset

End Property

Private Property Let YOriginOffset(ByVal lPiYels As Long)

    m_lYOriginOffset = lPiYels

End Property

Private Sub FindDirectory(ByVal sPath As String)

    If Len(Trim(sPath)) > 0 Then
        pvDisplayStartupPath sPath
    Else
        m_lSelectionID = 1
        m_lSelectedID = m_DisplayNodes(1)
    End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim lLast As Long
    Dim lMaxMove As Long
    Dim l As Long


    If Not bHasFocus Then Exit Sub

    lMaxMove = (llastdrawn - (lfirstdrawn + 1))
    m_bDrawingSelectedNode = False


    If KeyCode = 38 Then

        m_lSelectionID = m_lSelectionID - 1
        If m_lSelectionID <= LBound(m_DisplayNodes) + 1 Then m_lSelectionID = LBound(m_DisplayNodes) + 1
        m_lSelectedID = m_DisplayNodes(m_lSelectionID)
        pvRedraw

        'If m_Nodes(m_DisplayNodes(m_lSelectionID)).oRect.Top <= 1 Then
        If Not m_bDrawingSelectedNode Then m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) - 1
        'End If

    ElseIf KeyCode = 13 Then
        RaiseEvent NodeDblClick(m_lSelectionID, m_Nodes(m_DisplayNodes(m_lSelectionID)).Caption, m_Nodes(m_DisplayNodes(m_lSelectionID)).Path)
    ElseIf KeyCode = 39 Then

        l = m_lSelectionID
        KeyCode = 0

        If m_Nodes(m_DisplayNodes(m_lSelectionID)).State = ens_NotExpanded Then

            If m_lType = edt_Custom Then

                If m_lType = edt_Custom Then
                    Dim bExpanded As Boolean
                    Dim sCap As String
                    m_Nodes(m_DisplayNodes(m_lSelectionID)).Expanded = True
                    If m_Nodes(m_DisplayNodes(m_lSelectionID)).HasChildren Then
                        If Not m_Nodes(m_DisplayNodes(m_lSelectionID)).State = ens_ExpandedOnce Then
                            bExpanded = True
                            sCap = m_Nodes(m_DisplayNodes(m_lSelectionID)).Caption
                            RaiseEvent BeforeCustomNodeExpand(m_lSelectionID, sCap, bExpanded)
                            m_Nodes(m_DisplayNodes(m_lSelectionID)).State = ens_ExpandedOnce
                        End If
                    End If
                End If
            Else

                Screen.MousePointer = vbHourglass
                RaiseEvent Locating(m_Nodes(m_DisplayNodes(m_lSelectionID)).Path)

                DoEvents
                pvAddFolders m_DisplayNodes(m_lSelectionID)
                '-- Set The Expanded Flag
                m_Nodes(m_DisplayNodes(m_lSelectionID)).State = ens_ExpandedOnce
                UserControl.MousePointer = vbDefault
                m_Nodes(m_DisplayNodes(m_lSelectionID)).Expanded = True

                '-- Make the DisplayNodes
                pvMakeDisplayNodes

                If Not pbIsDimmed(hTable(m_Nodes(m_DisplayNodes(m_lSelectionID)).Children)) Then
                    m_Nodes(m_DisplayNodes(m_lSelectionID)).HasChildren = False
                Else
                    m_Nodes(m_DisplayNodes(m_lSelectionID)).ItemData = "(" & UBound(m_Nodes(m_DisplayNodes(m_lSelectionID)).Children) + 1 & ")"
                End If
                Screen.MousePointer = vbDefault
                RaiseEvent LocatingComplete

                pvRedraw
                Exit Sub
            End If
        Else
            m_Nodes(m_DisplayNodes(m_lSelectionID)).Expanded = True

            '-- Make the DisplayNodes
            pvMakeDisplayNodes

        End If
        pvRedraw
        Exit Sub
    ElseIf KeyCode = 37 Then
        '        KeyCode = 0
        Debug.Print m_lSelectedID
        l = m_lSelectionID

        m_Nodes(m_DisplayNodes(m_lSelectionID)).Expanded = False

        '-- Make the DisplayNodes
        pvMakeDisplayNodes
        pvRedraw
        'Exit Sub

    ElseIf KeyCode = 40 Then

        If m_lType = edt_Custom Then

            If m_lSelectedID + 1 >= UBound(m_DisplayNodes) Then

                m_lSelectedID = UBound(m_DisplayNodes)
                m_lSelectionID = UBound(m_DisplayNodes)

            Else


                m_lSelectionID = m_lSelectionID + 1
                If m_lSelectionID >= UBound(m_DisplayNodes) Then
                    m_lSelectionID = UBound(m_DisplayNodes)
                    m_cScrollBar.Value(efsVertical) = m_cScrollBar.Max(efsVertical)

                End If
                m_lSelectedID = m_DisplayNodes(m_lSelectionID)
                '

            End If
        Else
            m_lSelectionID = m_lSelectionID + 1
            If m_lSelectionID >= UBound(m_DisplayNodes) Then
                m_lSelectionID = UBound(m_DisplayNodes)
                m_cScrollBar.Value(efsVertical) = m_cScrollBar.Max(efsVertical)

            End If
            m_lSelectedID = m_DisplayNodes(m_lSelectionID)
            '

        End If
        pvRedraw
        If Not m_bDrawingSelectedNode Then m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) + 1

    ElseIf KeyCode = 34 Then

        m_lSelectionID = m_lSelectionID + lMaxMove
        If m_lSelectionID >= UBound(m_DisplayNodes) Then m_lSelectionID = UBound(m_DisplayNodes)
        m_lSelectedID = m_DisplayNodes(m_lSelectionID)
        pvRedraw

        If Not m_bDrawingSelectedNode Then m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) + lMaxMove



    ElseIf KeyCode = 33 Then


        m_lSelectionID = m_lSelectionID - lMaxMove
        If m_lSelectionID <= LBound(m_DisplayNodes) Then m_lSelectionID = LBound(m_DisplayNodes) + 1
        m_lSelectedID = m_DisplayNodes(m_lSelectionID)
        pvRedraw

        If Not m_bDrawingSelectedNode Then m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) - lMaxMove

    End If

    On Error GoTo 0

    DoEvents

    RaiseEvent SelectionChanged(m_Nodes(m_DisplayNodes(m_lSelectionID)).Caption, m_Nodes(m_DisplayNodes(m_lSelectionID)).Path)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'//---------------------------------------------------------------------------------------
' Procedure : GetFonts
' Type      : Sub
' DateTime  : 08/10/2004 12:51
' Author    : Gary Noble
' Purpose   : Loads All The Screen Fonts In To The Tree
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  08/10/2004
'//---------------------------------------------------------------------------------------
Private Sub GetFonts()

    Dim i As Long
    Dim lNode As Long
    Dim lFontNode As Long

    lNode = plAddNode("Available Fonts", "Available Fonts", -1, 0)
    m_Nodes(lNode).HasChildren = Screen.FontCount > 0
    m_Nodes(lNode).Expanded = True
    m_Nodes(lNode).State = ens_ExpandedOnce
    m_lSelectedID = 1

    For i = 1 To Screen.FontCount - 1
        lFontNode = plAddNode(Screen.Fonts(i), Screen.Fonts(i), -1, lNode)
        '-- Stops The System To Look For Directories
        m_Nodes(lFontNode).State = ens_ExpandedOnce
    Next
        m_Nodes(lNode).ItemData = "(" & UBound(m_Nodes(lNode).Children) & " Found)"
    
    '-- Note: Please Look A The pvRedraw Sub To Take Care Of The Drawing Of The Font Listings


End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvTVAdd
' Type      : Function
' DateTime  : 06/10/2004 17:36
' Author    : Gary Noble
' Purpose   : Global Calling Point For Adding a node To The array
' Returns   : Long
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  06/10/2004
'//---------------------------------------------------------------------------------------
Public Function CustomAdd( _
                 Optional ByVal hParent As Long = 0, _
                 Optional ByVal sText As String = vbNullString, _
                 Optional ByVal bForcePlusButton As Boolean = False _
                 , Optional ByVal sTag As String, Optional bExpanded As Boolean) As Long

    If m_lType < edt_Custom Then MsgBox "You Cannot Add To A Tree Thats Is Not Set As Custom", vbCritical, "xDirectory": Exit Function

    CustomAdd = pvTVAdd(hParent, , RTrim(pvStripNulls(sText)), , , , bForcePlusButton, sTag)

    m_Nodes(CustomAdd).Expanded = bExpanded


    If bExpanded Then
        m_Nodes(CustomAdd).HasChildren = bForcePlusButton
        If m_lType = edt_Custom Then
            m_Nodes(CustomAdd).Expanded = True
            If m_Nodes(CustomAdd).HasChildren Then
                If Not m_Nodes(CustomAdd).State = ens_ExpandedOnce Then
                    bExpanded = True
                    RaiseEvent BeforeCustomNodeExpand(CustomAdd, sText, bExpanded)
                    m_Nodes(m_DisplayNodes(CustomAdd)).State = ens_ExpandedOnce
                    m_Nodes(m_DisplayNodes(CustomAdd)).ItemData = "(" & UBound(m_Nodes(m_DisplayNodes(CustomAdd)).Children) + 1 & " Samples)"
                    
                    
                End If
            End If
        End If
        m_Nodes(CustomAdd).Expanded = bExpanded

    End If
End Function


Public Sub RefreshCustomNodes()
    If m_lType = edt_Custom Then pvMakeDisplayNodes: pvRedraw: pvMakeScrollScrollbarVisible
End Sub
