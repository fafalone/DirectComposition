Attribute VB_Name = "Module1"
Option Explicit
Public bEffectsActive As Boolean
Private Declare Function D3D11CreateDevice Lib "d3d11.dll" (ByVal pAdapter As IDXGIAdapter, ByVal DriverType As D3D_DRIVER_TYPE, ByVal Software As Long, ByVal flags As Long, pFeatureLevels As Any, ByVal featureLevels As Long, ByVal SDKVersion As Long, ppDevice As ID3D11Device, pFeatureLevel As D3D_FEATURE_LEVEL, ppImmediateContext As ID3D11DeviceContext) As Long

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (wndcls As WNDCLASSEX) As Integer
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal uFormat As DT_Flags) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As ShowWindowConstants) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As StockObjects) As Long
Private Declare Function AdjustWindowRect Lib "user32" (lpRect As RECT, ByVal dwStyle As WindowStyles, ByVal bMenu As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageW" (lpMsg As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (ByRef lpMsg As Any) As Long
Private Declare Function TranslateMessage Lib "user32" (ByRef lpMsg As Any) As Long
Private Declare Function CreateFontW Lib "gdi32" (ByVal cHeight As Long, ByVal cWidth As Long, ByVal cEscapement As Long, ByVal cOrientation As Long, ByVal cWeight As FontWeight, ByVal bItalic As Long, ByVal bUnderline As Long, ByVal bStrikeOut As Long, ByVal iCharset As FontCharset, ByVal iOutPrecision As FontOutPrecis, ByVal iClipPrecision As FontClipPrecis, ByVal iQuality As Long, ByVal iPitchAndFamily As Long, ByVal pszFaceName As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As SystemColors) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As LongPtr)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function UnregisterClassW Lib "user32" (ByVal lpClassName As LongPtr, Optional ByVal hInstance As LongPtr) As Long

Private Const ERROR_CLASS_ALREADY_EXISTS = &H582

Private Enum DefCursors
    IDC_ARROW = 32512&
    IDC_IBEAM = 32513&
    IDC_WAIT = 32514&
    IDC_CROSS = 32515&
    IDC_UPARROW = 32516&
    IDC_SIZE = 32640&
    IDC_ICON = 32641&
    IDC_SIZENWSE = 32642&
    IDC_SIZENESW = 32643&
    IDC_SIZEWE = 32644&
    IDC_SIZENS = 32645&
    IDC_SIZEALL = 32646&
    IDC_NO = 32648&
    IDC_HAND = 32649&
    IDC_APPSTARTING = 32650&
    IDC_HELP = 32651&
    IDC_PIN = 32671&
    IDC_PERSON = 32672&
End Enum

Private Enum SystemColors
    CTLCOLOR_MSGBOX = 0
    CTLCOLOR_EDIT = 1
    CTLCOLOR_LISTBOX = 2
    CTLCOLOR_BTN = 3
    CTLCOLOR_DLG = 4
    CTLCOLOR_SCROLLBAR = 5
    CTLCOLOR_STATIC = 6
    CTLCOLOR_MAX = 7

    COLOR_SCROLLBAR = 0
    COLOR_BACKGROUND = 1
    COLOR_ACTIVECAPTION = 2
    COLOR_INACTIVECAPTION = 3
    COLOR_MENU = 4
    COLOR_WINDOW = 5
    COLOR_WINDOWFRAME = 6
    COLOR_MENUTEXT = 7
    COLOR_WINDOWTEXT = 8
    COLOR_CAPTIONTEXT = 9
    COLOR_ACTIVEBORDER = 10
    COLOR_INACTIVEBORDER = 11
    COLOR_APPWORKSPACE = 12
    COLOR_HIGHLIGHT = 13
    COLOR_HIGHLIGHTTEXT = 14
    COLOR_BTNFACE = 15
    COLOR_BTNSHADOW = 16
    COLOR_GRAYTEXT = 17
    COLOR_BTNTEXT = 18
    COLOR_INACTIVECAPTIONTEXT = 19
    COLOR_BTNHIGHLIGHT = 20

    COLOR_3DDKSHADOW = 21
    COLOR_3DLIGHT = 22
    COLOR_INFOTEXT = 23
    COLOR_INFOBK = 24

    COLOR_HOTLIGHT = 26
    COLOR_GRADIENTACTIVECAPTION = 27
    COLOR_GRADIENTINACTIVECAPTION = 28

    COLOR_MENUHILIGHT = 29
    COLOR_MENUBAR = 30

    COLOR_DESKTOP = COLOR_BACKGROUND
    COLOR_3DFACE = COLOR_BTNFACE
    COLOR_3DSHADOW = COLOR_BTNSHADOW
    COLOR_3DHIGHLIGHT = COLOR_BTNHIGHLIGHT
    COLOR_3DHILIGHT = COLOR_BTNHIGHLIGHT
    COLOR_BTNHILIGHT = COLOR_BTNHIGHLIGHT
End Enum


Private Const CW_USEDEFAULT = &H80000000

Private Enum FontCharset
     ANSI_CHARSET = 0
     DEFAULT_CHARSET = 1
     SYMBOL_CHARSET = 2
     SHIFTJIS_CHARSET = 128
     HANGEUL_CHARSET = 129
     HANGUL_CHARSET = 129
     GB2312_CHARSET = 134
     CHINESEBIG5_CHARSET = 136
     OEM_CHARSET = 255
     JOHAB_CHARSET = 130
     HEBREW_CHARSET = 177
     ARABIC_CHARSET = 178
     GREEK_CHARSET = 161
     TURKISH_CHARSET = 162
     VIETNAMESE_CHARSET = 163
     THAI_CHARSET = 222
     EASTEUROPE_CHARSET = 238
     RUSSIAN_CHARSET = 204

     MAC_CHARSET = 77
     BALTIC_CHARSET = 186
End Enum

Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2
Private Const MONO_FONT = 8

Private Enum FontFamily
    FF_DONTCARE = 0
    FF_ROMAN = 16
    FF_SWISS = 32
    FF_MODERN = 48
    FF_SCRIPT = 64
    FF_DECORATIVE = 80
End Enum

Private Enum FontOutPrecis
    OUT_DEFAULT_PRECIS = 0
    OUT_STRING_PRECIS = 1
    OUT_CHARACTER_PRECIS = 2
    OUT_STROKE_PRECIS = 3
    OUT_TT_PRECIS = 4
    OUT_DEVICE_PRECIS = 5
    OUT_RASTER_PRECIS = 6
    OUT_TT_ONLY_PRECIS = 7
    OUT_OUTLINE_PRECIS = 8
    OUT_SCREEN_OUTLINE_PRECIS = 9
    OUT_PS_ONLY_PRECIS = 10
End Enum

Private Enum FontClipPrecis
    CLIP_DEFAULT_PRECIS = 0
    CLIP_CHARACTER_PRECIS = 1
    CLIP_STROKE_PRECIS = 2
    CLIP_MASK = &HF
'    CLIP_LH_ANGLES = (1 << 4)
'    CLIP_TT_ALWAYS = (2 << 4)
'    CLIP_DFA_DISABLE = (4 << 4)
'    CLIP_EMBEDDED = (8 << 4)
End Enum


Private Const DEFAULT_QUALITY As Byte = (0)
Private Const DRAFT_QUALITY As Byte = (1)
Private Const PROOF_QUALITY As Byte = (2)
Private Const NONANTIALIASED_QUALITY As Byte = (3)
Private Const ANTIALIASED_QUALITY As Byte = (4)
Private Const CLEARTYPE_QUALITY As Byte = (5)
Private Const CLEARTYPE_NATURAL_QUALITY As Byte = (6)

Private Enum FontWeight
    FW_DONTCARE = 0
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_HEAVY = 900
    FW_ULTRALIGHT = FW_EXTRALIGHT
    FW_REGULAR = FW_NORMAL
    FW_DEMIBOLD = FW_SEMIBOLD
    FW_ULTRABOLD = FW_EXTRABOLD
    FW_BLACK = FW_HEAVY
End Enum

Private Enum StockObjects
    WHITE_BRUSH = 0
    LTGRAY_BRUSH = 1
    GRAY_BRUSH = 2
    DKGRAY_BRUSH = 3
    BLACK_BRUSH = 4
    NULL_BRUSH = 5
    HOLLOW_BRUSH = NULL_BRUSH
    WHITE_PEN = 6
    BLACK_PEN = 7
    NULL_PEN = 8
    OEM_FIXED_FONT = 10
    ANSI_FIXED_FONT = 11
    ANSI_VAR_FONT = 12
    SYSTEM_FONT = 13
    DEVICE_DEFAULT_FONT = 14
    DEFAULT_PALETTE = 15
    SYSTEM_FIXED_FONT = 16
    DEFAULT_GUI_FONT = 17
    DC_BRUSH = 18
    DC_PEN = 19
End Enum

Private Enum ShowWindowConstants
     SW_HIDE = 0
     SW_SHOWNORMAL = 1
     SW_NORMAL = 1
     SW_SHOWMINIMIZED = 2
     SW_SHOWMAXIMIZED = 3
     SW_MAXIMIZE = 3
     SW_SHOWNOACTIVATE = 4
     SW_SHOW = 5
     SW_MINIMIZE = 6
     SW_SHOWMINNOACTIVE = 7
     SW_SHOWNA = 8
     SW_RESTORE = 9
     SW_SHOWDEFAULT = 10
     SW_FORCEMINIMIZE = 11
     SW_MAX = 11
End Enum

Private Enum DT_Flags
    DT_BOTTOM = &H8&
    DT_CENTER = &H1&
    DT_LEFT = &H0&
    DT_CALCRECT = &H400&
    DT_WORDBREAK = &H10&
    DT_VCENTER = &H4&
    DT_TOP = &H0&
    DT_TABSTOP = &H80&
    DT_SINGLELINE = &H20&
    DT_RIGHT = &H2&
    DT_NOCLIP = &H100&
    DT_INTERNAL = &H1000&
    DT_EXTERNALLEADING = &H200&
    DT_EXPANDTABS = &H40&
    DT_CHARSTREAM = 4&
    DT_NOPREFIX = &H800&
    DT_EDITCONTROL = &H2000&
    DT_PATH_ELLIPSIS = &H4000&
    DT_END_ELLIPSIS = &H8000&
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
End Enum

Private Const TRANSPARENT = 1&
Private Const OPAQUE = 2&

Private Const CFalse As Long = 0
Private Const CTrue As Long = 1

Private bpTrigger As Boolean


Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(0 To 31) As Byte
End Type


Private Enum ClassStyles
     CS_VREDRAW = &H1
     CS_HREDRAW = &H2
     CS_DBLCLKS = &H8
     CS_OWNDC = &H20
     CS_CLASSDC = &H40
     CS_PARENTDC = &H80
     CS_NOCLOSE = &H200
     CS_SAVEBITS = &H800
     CS_BYTEALIGNCLIENT = &H1000
     CS_BYTEALIGNWINDOW = &H2000
     CS_GLOBALCLASS = &H4000
     CS_IME = &H10000
     CS_DROPSHADOW = &H20000
End Enum
Private Type WNDCLASSEX
    cbSize As Long
    style As ClassStyles
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As Long
    hIconSm As Long
End Type

#If Win64 Then
    Private Const cbPtr As Long = 8
#Else
    Private Const cbPtr As Long = 4
#End If

Private Const D3D11_SDK_VERSION = 7


Private Const FONT_TYPEFACE = "Segoe UI Light"
Private fontTypeface(31) As Integer

Private Const FONT_HEIGHT_LOGO = 0
Private fontHeightLogo As Long
Private Const FONT_HEIGHT_TITLE = 50
Private fontHeightTitle As Long
Private Const FONT_HEIGHT_DESCRIPTION = 22
Private fontHeightDescription As Long

Private m_hWnd As Long

Private tileSize As Long

Private windowWidth As Long
Private windowHeight As Long

Private d3d11Device As ID3D11Device
Private d3d11DeviceContext As ID3D11DeviceContext

Private d2d1Factory As ID2D1Factory1

Private d2d1Device As ID2D1Device
Private d2d1DeviceContext As ID2D1DeviceContext

Private device As IDCompositionDevice
Private target As IDCompositionTarget
Private visual As IDCompositionVisual
Private visualLeft As IDCompositionVisual
Private visualLeftChild(3) As IDCompositionVisual
Private visualRight As IDCompositionVisual

Private surfaceLeftChild(3) As IDCompositionSurface

Private effectGroupLeft As IDCompositionEffectGroup
Private effectGroupLeftChild(3) As IDCompositionEffectGroup
Private effectGroupRight As IDCompositionEffectGroup

Private currentVisual As Long

Private Enum VIEW_STATE
    ZOOMEDOUT
    ZOOMEDIN
End Enum

Private state As VIEW_STATE

Private Enum ACTION_TYPE
    ZoomOut
    ZoomIn
End Enum

Private actionType As ACTION_TYPE

Private Const wndClass = "MainWindowClass"
Private Const wndName = "DirectComposition Effects Sample"

Private Const gridSize = 100
Public Sub PostLog(sMsg As String)
Debug.Print sMsg
Exit Sub
End Sub

Private Function SUCCEEDED(hr As Long) As Boolean
    SUCCEEDED = (hr >= 0)
End Function
Sub Main()
    RunDCompEffect
    
End Sub

Public Function RunDCompEffect() As Long
    bEffectsActive = True
    tileSize = (3 * gridSize)
    windowWidth = (9 * gridSize)
    windowHeight = (6 * gridSize)
    state = (ZOOMEDOUT)
    actionType = (ZoomOut)
    currentVisual = (0)

        Dim result As Long
    
    If SUCCEEDED(BeforeEnteringMessageLoop()) Then
        PostLog "BeforeEnteringMessageLoop succeeded"
        result = EnterMessageLoop()
    Else
        MsgBox "An error occured when running the sample.", vbCritical Or vbOKOnly, wndName
    End If
    
    AfterLeavingMessageLoop
    RunDCompEffect = result
End Function

Private Function BeforeEnteringMessageLoop() As Long
    Dim hr As Long
    hr = CreateApplicationWindow()
    If SUCCEEDED(hr) Then
        hr = CreateD3D11Device()
    Else
        PostLog "CreateApplicationWindow failed, hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        hr = CreateD2D1Factory()
    Else
        PostLog "CreateD3D11Device failed, hr=0x" & Hex$(hr)
    End If
        
    If SUCCEEDED(hr) Then
        hr = CreateD2D1Device()
    Else
        PostLog "CreateD2D1Factory failed, hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        hr = CreateDCompositionDevice()
    Else
        PostLog "CreateD2D1Device failed, hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        hr = CreateDCompositionVisualTree()
    Else
        PostLog "CreateDCompositionDevice failed, hr=0x" & Hex$(hr)
    End If

    If SUCCEEDED(hr) Then
        PostLog "BeforeEnteringMessageLoop->All initialization routines returned success."
    Else
        PostLog "CreateDCompositionVisualTree failed, hr=0x" & Hex$(hr)
    End If
    BeforeEnteringMessageLoop = hr
End Function

Private Function EnterMessageLoop() As Long
    Dim result As Long
    
    If ShowApplicationWindow() Then
        Dim tMSG As MSG
        Dim hr As Long
        PostLog "Entering message loop"
        hr = GetMessage(tMSG, 0, 0, 0)
        Do While hr <> 0
            If hr = -1 Then
                PostLog "Error: 0x" & Hex$(Err.LastDllError)
            Else
                TranslateMessage tMSG
                DispatchMessage tMSG
            End If
            hr = GetMessage(tMSG, 0, 0, 0)
        Loop
        PostLog "Exited message loop"
        result = tMSG.wParam
    End If

    
    EnterMessageLoop = result
End Function

Private Sub AfterLeavingMessageLoop()

     DestroyDCompositionVisualTree
    
     DestroyDCompositionDevice

     DestroyD2D1Device

     DestroyD2D1Factory

     DestroyD3D11Device

    DestroyApplicationWindow
End Sub

Private Function CreateApplicationWindow() As Long

    Dim hr   As Long: hr = S_OK
    
    Dim wcex As WNDCLASSEX
    
    wcex.cbSize = LenB(wcex)
    wcex.style = CS_HREDRAW Or CS_VREDRAW
    wcex.lpfnWndProc = FARPROC(AddressOf WindowProc)
    wcex.cbClsExtra = 0
    wcex.cbWndExtra = 0
    wcex.hInstance = App.hInstance
    wcex.hIcon = 0
    wcex.hCursor = LoadCursor(0, IDC_ARROW)
    wcex.hbrBackground = GetStockObject(WHITE_BRUSH)
    wcex.lpszMenuName = 0
    wcex.lpszClassName = StrPtr(wndClass)
    wcex.hIconSm = 0
    
    hr = IIf(RegisterClassEx(wcex), S_OK, E_FAIL)
    If Err.LastDllError = ERROR_CLASS_ALREADY_EXISTS Then
        PostLog "ERROR_CLASS_ALREADY_EXISTS; registering."
        UnregisterClassW StrPtr(wndClass), App.hInstance
        hr = IIf(RegisterClassEx(wcex), S_OK, E_FAIL)
    End If
    
    If SUCCEEDED(hr) Then
        PostLog "RegisterClassEx succeeded"

        Dim RECT As RECT

        RECT.Right = windowWidth: RECT.Bottom = windowHeight
        
        AdjustWindowRect RECT, WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_MINIMIZEBOX, 0
        
        m_hWnd = CreateWindowExW(0, StrPtr(wndClass), StrPtr(wndName), WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_MINIMIZEBOX, CW_USEDEFAULT, CW_USEDEFAULT, RECT.Right - RECT.Left, RECT.Bottom - RECT.Top, 0, 0, App.hInstance, ByVal 0)
        
        PostLog "Window hwnd=" & m_hWnd

        If m_hWnd = 0 Then
            hr = E_UNEXPECTED
            PostLog "CreateWindowExW failed, LastError=0x" & Hex$(Err.LastDllError)
        End If

    Else
        PostLog "RegisterClassEx failed, LastError=0x" & Hex$(Err.LastDllError)
    End If
    
    If SUCCEEDED(hr) Then
        'CopyMemory fontTypeface(0), ByVal StrPtr(FONT_TYPEFACE), LenB(FONT_TYPEFACE)
        
        fontHeightLogo = FONT_HEIGHT_LOGO
        
        fontHeightTitle = FONT_HEIGHT_TITLE
        
        fontHeightDescription = FONT_HEIGHT_DESCRIPTION
    End If
    
    CreateApplicationWindow = hr
             
End Function
Private Function FARPROC(lpfn As Long) As Long
FARPROC = lpfn
End Function
Private Function ShowApplicationWindow() As Boolean
    Dim bSucceeded As Boolean: bSucceeded = (m_hWnd <> 0)
    PostLog "ShowApplicationWindow bSucceeded=" & bSucceeded
    If bSucceeded Then
        ShowWindow m_hWnd, SW_SHOW
        UpdateWindow m_hWnd
    End If
    ShowApplicationWindow = bSucceeded
End Function

Private Sub DestroyApplicationWindow()
    If m_hWnd Then
        DestroyWindow m_hWnd
        UnregisterClassW StrPtr(wndClass), App.hInstance
    End If
End Sub

Private Sub DestroyDCompositionVisualTree()
    
    Set effectGroupRight = Nothing
    
    Dim i As Long
    
    For i = 0 To 3
        Set effectGroupLeftChild(i) = Nothing
    Next
    
    For i = 0 To 3
        Set surfaceLeftChild(i) = Nothing
    Next
    
    Set visualRight = Nothing
    
    For i = 0 To 3
        Set visualLeftChild(i) = Nothing
    Next
    
    Set visualLeft = Nothing
    
    Set visual = Nothing
    
    Set target = Nothing
End Sub

Private Sub DestroyDCompositionDevice()
    Set device = Nothing
End Sub

Private Sub DestroyD3D11Device()
    Set d3d11DeviceContext = Nothing
    Set d3d11Device = Nothing
End Sub

Private Sub DestroyD2D1Device()
    Set d2d1DeviceContext = Nothing
    Set d2d1Device = Nothing
End Sub

Private Sub DestroyD2D1Factory()
    Set d2d1Factory = Nothing
End Sub

Private Function CreateD3D11Device() As Long
    PostLog "CreateD3D11Device->Entry"
    Dim hr As Long '= S_OK
    Dim i As Long
    Dim driverTypes(1) As D3D_DRIVER_TYPE
    driverTypes(0) = D3D_DRIVER_TYPE_HARDWARE
    driverTypes(1) = D3D_DRIVER_TYPE_WARP
    
    Dim featureLevelSupported As D3D_FEATURE_LEVEL
    
    
    For i = 0 To UBound(driverTypes)
        
        hr = D3D11CreateDevice(Nothing, driverTypes(i), 0, D3D11_CREATE_DEVICE_BGRA_SUPPORT, ByVal 0, 0, D3D11_SDK_VERSION, d3d11Device, featureLevelSupported, d3d11DeviceContext)
        
        If SUCCEEDED(hr) Then
            PostLog "D3D11CreateDevice(" & i & ") succeeded."
            Exit For
        End If
    Next
    PostLog "CreateD3D11Device->hr=0x" & Hex$(hr)
    CreateD3D11Device = hr
            
End Function

Private Function CreateD2D1Factory() As Long
    PostLog "CreateD2D1Factory->Entry"
    Set d2d1Factory = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, IID_ID2D1Factory, ByVal 0)
    If (d2d1Factory Is Nothing) Then
        CreateD2D1Factory = E_NOINTERFACE
    End If
End Function

Private Function CreateD2D1Device() As Long
    PostLog "CreateD2D1Device->Entry"
    On Error GoTo e0
    Dim hr As Long: hr = IIf((d3d11Device Is Nothing) Or (d2d1Factory Is Nothing), E_UNEXPECTED, S_OK)
    
    Dim dxgiDevice As IDXGIDevice
    
    If SUCCEEDED(hr) Then
        Set dxgiDevice = d3d11Device
        'hr = Err.LastHResult '(dxgiDevice IsNot Nothing)
    End If
    If dxgiDevice Is Nothing Then
        hr = E_NOINTERFACE
    End If
    If SUCCEEDED(hr) Then
        PostLog "CreateD2D1Device->Successful QI for dxgiDevice"
        d2d1Factory.CreateDevice dxgiDevice, d2d1Device
        'hr = Err.LastHResult
    Else
        PostLog "CreateD2D1Device->Failed QI for dxgiDevice, lasthr=0x" & Hex$(hr) 'Hex$(Err.LastHResult)
    End If
    
    If SUCCEEDED(hr) Then
        PostLog "CreateD2D1Device->Created d2d1device"
        d2d1Device.CreateDeviceContext D2D1_DEVICE_CONTEXT_OPTIONS_NONE, d2d1DeviceContext
        'hr = Err.LastHResult
    Else
        PostLog "CreateD2D1Device->Failed to create d2d1device, lasthr=0x" & Hex$(hr)
    End If
    
    CreateD2D1Device = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks
End Function


Private Function CreateDCompositionDevice() As Long
    PostLog "CreateDCompositionDevice->Entry"
    Dim hr As Long: hr = IIf(d3d11Device Is Nothing, E_UNEXPECTED, S_OK)
    
    Dim dxgiDevice As IDXGIDevice
    
    If SUCCEEDED(hr) Then
        Dim pUnk As oleexp.IUnknown 'IUnknownUnrestricted
        Set pUnk = d3d11Device
        hr = pUnk.QueryInterface(IID_IDXGIDevice, dxgiDevice)
        PostLog "CreateDCompositionDevice->d3d11device ok, dxgiDevice QI hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        hr = DCompositionCreateDevice(dxgiDevice, IID_IDCompositionDevice, device)
    End If
    
    CreateDCompositionDevice = hr
End Function

Private Function CreateDCompositionVisualTree() As Long
    PostLog "CreateDCompositionVisualTree->Entry"
    On Error GoTo e0
    Dim hr As Long: hr = IIf(device Is Nothing, E_UNEXPECTED, S_OK)
    
    If SUCCEEDED(hr) Then
        device.CreateVisual visual
        'hr = Err.LastHResult
    End If
    PostLog "CreateDCompositionVisualTree->visual hr=0x" & Hex$(hr)
    If SUCCEEDED(hr) Then
        device.CreateVisual visualLeft
        'hr = Err.LastHResult
    End If
    PostLog "CreateDCompositionVisualTree->visualLeft hr=0x" & Hex$(hr)
    Dim surfaceLeft As IDCompositionSurface
    
    If SUCCEEDED(hr) Then
        hr = CreateSurface(tileSize, 1#, 0, 0, surfaceLeft)
        PostLog "CreateDCompositionVisualTree->CreateSurface(surfaceLeft) hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        visualLeft.SetContent surfaceLeft
        'hr = Err.LastHResult
        PostLog "CreateDCompositionVisualTree->SetContent(surfaceLeft) hr=0x" & Hex$(hr)
    End If
    
    Dim i As Long
    
    For i = 0 To 3
        
        If SUCCEEDED(hr) Then
            device.CreateVisual visualLeftChild(i)
            'hr = Err.LastHResult
        End If
        
        If i = 0 Then
            If SUCCEEDED(hr) Then
                hr = CreateSurface(tileSize, 0, 1#, 0, surfaceLeftChild(i))
            Else
                PostLog "CreateDCompositionVisualTree->device.CreateVisual(VisualLeftChild(" & i & ")) failed, hr=0x" & Hex$(hr)
            End If
            
        ElseIf i = 1 Then
            If SUCCEEDED(hr) Then
                hr = CreateSurface(tileSize, 0.5, 0, 0.5, surfaceLeftChild(i))
            Else
                PostLog "CreateDCompositionVisualTree->device.CreateVisual(VisualLeftChild(" & i & ")) failed, hr=0x" & Hex$(hr)
            End If
        
        ElseIf i = 2 Then
            If SUCCEEDED(hr) Then
                hr = CreateSurface(tileSize, 0.5, 0.5, 0, surfaceLeftChild(i))
            Else
                PostLog "CreateDCompositionVisualTree->device.CreateVisual(VisualLeftChild(" & i & ")) failed, hr=0x" & Hex$(hr)
            End If
            
        ElseIf i = 3 Then
            If SUCCEEDED(hr) Then
                hr = CreateSurface(tileSize, 0, 0, 1#, surfaceLeftChild(i))
            Else
                PostLog "CreateDCompositionVisualTree->device.CreateVisual(VisualLeftChild(" & i & ")) failed, hr=0x" & Hex$(hr)
            End If
        End If
        
        If SUCCEEDED(hr) Then
            visualLeftChild(i).SetContent surfaceLeftChild(i)
            'hr = Err.LastHResult
        Else
            PostLog "CreateDCompositionVisualTree->CreateSurface(surfaceLeftChild(" & i & ")) fail; hr=0x" & Hex$(hr)
        End If
    Next
    
    If SUCCEEDED(hr) Then
        device.CreateVisual visualRight
        'hr = Err.LastHResult
        PostLog "CreateDCompositionVisualTree->CreateVisual(visualRight) hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        visualRight.SetContent surfaceLeftChild(currentVisual)
        'hr = Err.LastHResult
        PostLog "CreateDCompositionVisualTree->SetContent surfaceLeftChild(cv) hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        visual.AddVisual visualLeft, 1, Nothing
        'hr = Err.LastHResult
        PostLog "CreateDCompositionVisualTree->AddVisual visualLeft hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        For i = 0 To 3
            visualLeft.AddVisual visualLeftChild(i), 0, Nothing
            'hr = Err.LastHResult
        Next
        PostLog "CreateDCompositionVisualTree->visualLeftChildren i=" & i & ", hr=0x" & Hex$(hr)
    End If
    
    
    If SUCCEEDED(hr) Then
        visual.AddVisual visualRight, 1, visualLeft
        'hr = Err.LastHResult
        PostLog "CreateDCompositionVisualTree->AddVisual visualRight,visualLeft hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        hr = SetEffectOnVisuals()
    End If
    
    If SUCCEEDED(hr) Then
        device.CreateTargetForHwnd m_hWnd, 1, target
        'hr = Err.LastHResult
        PostLog "CreateDCompositionVisualTree->CreateTargetForHwnd hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        target.SetRoot visual
        'hr = Err.LastHResult
        PostLog "CreateDCompositionVisualTree->SetRoot hr=0x" & Hex$(hr)
    End If
    
    If SUCCEEDED(hr) Then
        device.Commit
        'hr = Err.LastHResult
    End If
    
    CreateDCompositionVisualTree = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks
End Function

Private Function SetEffectOnVisuals() As Long
    Dim hr As Long: hr = SetEffectOnVisualLeft()
    
    If SUCCEEDED(hr) Then
        hr = SetEffectOnVisualLeftChildren()
    End If
    
    If SUCCEEDED(hr) Then
        hr = SetEffectOnVisualRight()
    End If
    
    SetEffectOnVisuals = hr
End Function

Private Function SetEffectOnVisualLeftChildren() As Long
    PostLog "SetEffectOnVisualLeftChildren->Entry"
    On Error GoTo e0
   
    Dim hr As Long ' = S_OK
    
    Dim i As Long
    For i = 0 To 3
        Dim r As Long: r = i \ 2
        Dim c As Long: c = i Mod 2
        
        Dim oscale As IDCompositionScaleTransform3D
        
        If SUCCEEDED(hr) Then
            hr = CreateScaleTransform( _
                0#, 0#, 0#, _
                1# / 3#, 1# / 3#, 1#, _
                oscale)
        End If
        
        Dim translate As IDCompositionTranslateTransform3D
        
        If SUCCEEDED(hr) Then
            hr = CreateTranslateTransform((0.25 + c * 1.5) * gridSize, (0.25 + r * 1.5) * gridSize, 0#, translate)
        End If
        
        Dim transforms(0 To 1) As IDCompositionTransform3D
        Set transforms(0) = oscale
        Set transforms(1) = translate
        
        Dim transformGroup As IDCompositionTransform3D
        
        If SUCCEEDED(hr) Then
            device.CreateTransform3DGroup VarPtr(transforms(0)), 2, transformGroup
            'hr = Err.LastHResult
        End If
        
        If SUCCEEDED(hr) Then
            Set effectGroupLeftChild(i) = Nothing
            device.CreateEffectGroup effectGroupLeftChild(i)
            'hr = Err.LastHResult
        End If
        
        If SUCCEEDED(hr) Then
            effectGroupLeftChild(i).SetTransform3D transformGroup
            'hr = Err.LastHResult
        End If
        
        If SUCCEEDED(hr) And (i = currentVisual) Then
            Dim opacityAnimation As IDCompositionAnimation
            
            Dim beginOpacity As Single: beginOpacity = IIf(actionType = ZoomOut, 1#, 0#)
            Dim endOpacity As Single: endOpacity = IIf(actionType = ZoomOut, 0#, 1#)
            
            hr = CreateLinearAnimation(beginOpacity, endOpacity, 0.25, 1.25, opacityAnimation)
            
            If SUCCEEDED(hr) Then
                effectGroupLeftChild(i).SetOpacity_A opacityAnimation
                'hr = Err.LastHResult
            End If
        End If
        
        If SUCCEEDED(hr) Then
            visualLeftChild(i).SetEffect effectGroupLeftChild(i)
            'hr = Err.LastHResult
        End If
        
    Next
    
    SetEffectOnVisualLeftChildren = hr
Exit Function
e0:
hr = Err.Number
Resume Next

End Function

Private Function CreateScaleTransform(centerX As Single, centerY As Single, centerZ As Single, scaleX As Single, scaleY As Single, scaleZ As Single, scaleTransform As IDCompositionScaleTransform3D) As Long
    PostLog "CreateScaleTransform->Entry"
    On Error GoTo e0
     Dim hr As Long: hr = IIf(VarPtr(scaleTransform) = 0, E_POINTER, S_OK)
    
    If SUCCEEDED(hr) Then
        Set scaleTransform = Nothing
        hr = IIf(device Is Nothing, E_UNEXPECTED, S_OK)
    End If
    
    Dim transform As IDCompositionScaleTransform3D
    
    If SUCCEEDED(hr) Then
        device.CreateScaleTransform3D transform
    End If
    
    
    If SUCCEEDED(hr) Then
        transform.SetCenterX centerX
        'hr = Err.LastHResult
    End If
    
     If SUCCEEDED(hr) Then
        transform.SetCenterY centerY
        'hr = Err.LastHResult
    End If

    If SUCCEEDED(hr) Then
        transform.SetCenterZ centerZ
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        transform.SetScaleX scaleX
        'hr = Err.LastHResult
    End If
    
     If SUCCEEDED(hr) Then
        transform.SetScaleY scaleY
        'hr = Err.LastHResult
    End If

    If SUCCEEDED(hr) Then
        transform.SetScaleZ scaleZ
        'hr = Err.LastHResult
    End If

    If SUCCEEDED(hr) Then
        CopyMemory scaleTransform, transform, cbPtr
        ZeroMemory transform, cbPtr
    End If

    CreateScaleTransform = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks
End Function

Private Function SetEffectOnVisualRight() As Long
    PostLog "SetEffectOnVisualRight->Entry"
    On Error GoTo e0
    Dim hr As Long: hr = IIf(visualRight Is Nothing, E_UNEXPECTED, S_OK)
    
    Dim beginOffsetX As Single: beginOffsetX = IIf(actionType = ZoomOut, 6.5, 3.75)
    Dim endOffsetX As Single: endOffsetX = IIf(actionType = ZoomOut, 3.75, 6.5)
    Dim offsetY As Single: offsetY = 1.5
    
    Dim translateTransform As IDCompositionTranslateTransform3D
    
    If SUCCEEDED(hr) Then
        hr = CreateTranslateTransform_A(beginOffsetX * gridSize, offsetY * gridSize, 0, endOffsetX * gridSize, offsetY * gridSize, 0, 0.25, 1.25, translateTransform)
    End If
    
    Dim transforms(0) As IDCompositionTransform3D
    Set transforms(0) = translateTransform
    
    Dim transformGroup As IDCompositionTransform3D
    
    If SUCCEEDED(hr) Then
        device.CreateTransform3DGroup VarPtr(transforms(0)), 1, transformGroup
        'hr = Err.LastHResult
    End If

    If SUCCEEDED(hr) Then
        Set effectGroupRight = Nothing
        device.CreateEffectGroup effectGroupRight
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        effectGroupRight.SetTransform3D transformGroup
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        Dim opacityAnimation As IDCompositionAnimation
        
        Dim beginOpacity As Single: beginOpacity = IIf(actionType = ZoomOut, 0#, 1#)
        Dim endOpacity As Single: endOpacity = IIf(actionType = ZoomOut, 1#, 0#)
        
        hr = CreateLinearAnimation(beginOpacity, endOpacity, 0.25, 1.25, opacityAnimation)
        
        If SUCCEEDED(hr) Then
            effectGroupRight.SetOpacity_A opacityAnimation
            'hr = Err.LastHResult
        End If
    End If
    
    If SUCCEEDED(hr) Then
        visualRight.SetEffect effectGroupRight
        'hr = Err.LastHResult
    End If
    
    SetEffectOnVisualRight = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks
End Function

Private Function SetEffectOnVisualLeft() As Long
    PostLog "SetEffectOnVisualLeft->Entry"
    On Error GoTo e0
    Dim hr As Long: hr = IIf(visualLeft Is Nothing, E_UNEXPECTED, S_OK)
    
    Dim beginOffsetX As Single: beginOffsetX = IIf(actionType = ZoomOut, 3#, 0.5)
    Dim endOffsetX As Single: endOffsetX = IIf(actionType = ZoomOut, 0.5, 3#)
    Dim offsetY As Single: offsetY = 1.5
    
    Dim beginAngle As Single: beginAngle = IIf(actionType = ZoomOut, 0, 30)
    Dim endAngle As Single: endAngle = IIf(actionType = ZoomOut, 30, 0)
    
    Dim translateTransform As IDCompositionTranslateTransform3D
    
    If SUCCEEDED(hr) Then
        hr = CreateTranslateTransform_A(beginOffsetX * gridSize, offsetY * gridSize, 0, endOffsetX * gridSize, offsetY * gridSize, 0, 0.25, 1.25, translateTransform)
    End If
    
    Dim rotateTransform As IDCompositionRotateTransform3D

    If SUCCEEDED(hr) Then
        hr = CreateRotateTransform(3.5 * gridSize, 1.5 * gridSize, 0, 0, 1#, 0, beginAngle, endAngle, 0.25, 1.25, rotateTransform)
    End If
    
    Dim perspectiveTransform As IDCompositionMatrixTransform3D
    
    If SUCCEEDED(hr) Then
        hr = CreatePerspectiveTransform(0#, 0#, -1# / (9# * gridSize), perspectiveTransform)
    End If
    
    Dim transforms(0 To 2) As IDCompositionTransform3D
    Set transforms(0) = translateTransform
    Set transforms(1) = rotateTransform
    Set transforms(2) = perspectiveTransform
    
    Dim transformGroup As IDCompositionTransform3D
    
    If SUCCEEDED(hr) Then
        device.CreateTransform3DGroup VarPtr(transforms(0)), 3, transformGroup
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        Set effectGroupLeft = Nothing
        device.CreateEffectGroup effectGroupLeft
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        effectGroupLeft.SetTransform3D transformGroup
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        visualLeft.SetEffect effectGroupLeft
        'hr = Err.LastHResult
    End If
    
    SetEffectOnVisualLeft = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks
End Function

Private Function CreatePerspectiveTransform(dx As Single, dy As Single, dz As Single, perspectiveTransform As IDCompositionMatrixTransform3D) As Long
    PostLog "CreatePerspectiveTransform->Entry"
    On Error GoTo e0
    Dim hr As Long: hr = IIf(VarPtr(perspectiveTransform) = 0, E_POINTER, S_OK)
    
    If SUCCEEDED(hr) Then
        Set perspectiveTransform = Nothing
    End If
    
    'VB lays out matrices different- we swap the rows and columns
    Dim matrix As D3DMATRIX
    matrix.m(0, 0) = 1#: matrix.m(1, 0) = 0#: matrix.m(2, 0) = 0#: matrix.m(3, 0) = dx
    matrix.m(0, 1) = 0#: matrix.m(1, 1) = 1#: matrix.m(2, 1) = 0#: matrix.m(3, 1) = dy
    matrix.m(0, 2) = 0#: matrix.m(1, 2) = 0#: matrix.m(2, 2) = 1#: matrix.m(3, 2) = dz
    matrix.m(0, 3) = 0#: matrix.m(1, 3) = 0#: matrix.m(2, 3) = 0#: matrix.m(3, 3) = 1#
    
    Dim transform As IDCompositionMatrixTransform3D
    
    If SUCCEEDED(hr) Then
        device.CreateMatrixTransform3D transform
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        transform.SetMatrix matrix
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        CopyMemory perspectiveTransform, transform, cbPtr
        ZeroMemory transform, cbPtr
    End If

    CreatePerspectiveTransform = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks
        
End Function

 
Private Function CreateRotateTransform(centerX As Single, centerY As Single, centerZ As Single, _
                                       axisX As Single, axisY As Single, axisZ As Single, _
                                       beginAngle As Single, endAngle As Single, beginTime As Single, endTime As Single, _
                                       rotateTransform As IDCompositionRotateTransform3D) As Long
                                           
    PostLog "CreateRotateTransform->Entry"
    On Error GoTo e0
     Dim hr As Long: hr = IIf(VarPtr(rotateTransform) = 0, E_POINTER, S_OK)
    
    If SUCCEEDED(hr) Then
        Set rotateTransform = Nothing
        hr = IIf(device Is Nothing, E_UNEXPECTED, S_OK)
    End If
    
    Dim transform As IDCompositionRotateTransform3D
    
    If SUCCEEDED(hr) Then
        device.CreateRotateTransform3D transform
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        transform.SetCenterX centerX
        'hr = Err.LastHResult
    End If
    
     If SUCCEEDED(hr) Then
        transform.SetCenterY centerY
        'hr = Err.LastHResult
    End If

    If SUCCEEDED(hr) Then
        transform.SetCenterZ centerZ
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        transform.SetAxisX axisX
        'hr = Err.LastHResult
    End If
    
     If SUCCEEDED(hr) Then
        transform.SetAxisY axisY
        'hr = Err.LastHResult
    End If

    If SUCCEEDED(hr) Then
        transform.SetAxisZ axisZ
        'hr = Err.LastHResult
    End If

    Dim angleAnimation As IDCompositionAnimation
    
    If SUCCEEDED(hr) Then
        hr = CreateLinearAnimation(beginAngle, endAngle, beginTime, endTime, angleAnimation)
    End If
    
    If SUCCEEDED(hr) Then
        transform.SetAngle_A angleAnimation
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        CopyMemory rotateTransform, transform, cbPtr
        ZeroMemory transform, cbPtr
    End If

    CreateRotateTransform = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks
        
        
End Function

Private Function CreateTranslateTransform(offsetX As Single, offsetY As Single, offsetZ As Single, translateTransform As IDCompositionTranslateTransform3D) As Long
    PostLog "CreateTranslateTransform->Entry"
    On Error GoTo e0
    Dim hr As Long: hr = IIf(VarPtr(translateTransform) = 0, E_POINTER, S_OK)
    
    If SUCCEEDED(hr) Then
        Set translateTransform = Nothing
        hr = IIf(device Is Nothing, E_UNEXPECTED, S_OK)
    End If
    
    Dim transform As IDCompositionTranslateTransform3D
    
    If SUCCEEDED(hr) Then
        device.CreateTranslateTransform3D transform
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        transform.SetOffsetX offsetX
        'hr = Err.LastHResult
    End If

    If SUCCEEDED(hr) Then
        transform.SetOffsetY offsetY
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        transform.SetOffsetZ offsetZ
        'hr = Err.LastHResult
    End If

    If SUCCEEDED(hr) Then
        CopyMemory translateTransform, transform, cbPtr
        ZeroMemory transform, cbPtr
    End If

    CreateTranslateTransform = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks
End Function

                                       
Private Function CreateTranslateTransform_A(beginOffsetX As Single, beginOffsetY As Single, beginOffsetZ As Single, endOffsetX As Single, endOffsetY As Single, endOffsetZ As Single, beginTime As Single, endTime As Single, translateTransform As IDCompositionTranslateTransform3D) As Long
    PostLog "CreateTranslateTransform_A->Entry"
    On Error GoTo e0

    Dim hr As Long: hr = IIf(VarPtr(translateTransform) = 0, E_POINTER, S_OK)
    
    If SUCCEEDED(hr) Then
        Set translateTransform = Nothing
        hr = IIf(device Is Nothing, E_UNEXPECTED, S_OK)
    End If
    
    Dim transform As IDCompositionTranslateTransform3D
    
    If SUCCEEDED(hr) Then
        device.CreateTranslateTransform3D transform
        'hr = Err.LastHResult
    End If
    
    Dim offsetXAnimation As IDCompositionAnimation
    
    If SUCCEEDED(hr) Then
        hr = CreateLinearAnimation(beginOffsetX, endOffsetX, beginTime, endTime, offsetXAnimation)
    End If
    
    If SUCCEEDED(hr) Then
        transform.SetOffsetX_A offsetXAnimation
        'hr = Err.LastHResult
    End If

    Dim offsetYAnimation As IDCompositionAnimation
    
    If SUCCEEDED(hr) Then
        hr = CreateLinearAnimation(beginOffsetY, endOffsetY, beginTime, endTime, offsetYAnimation)
    End If
    
    If SUCCEEDED(hr) Then
        transform.SetOffsetY_A offsetYAnimation
        'hr = Err.LastHResult
    End If
    
    Dim offsetZAnimation As IDCompositionAnimation
    
    If SUCCEEDED(hr) Then
        hr = CreateLinearAnimation(beginOffsetZ, endOffsetZ, beginTime, endTime, offsetZAnimation)
    End If
    
    If SUCCEEDED(hr) Then
        transform.SetOffsetZ_A offsetZAnimation
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        CopyMemory translateTransform, transform, cbPtr
        ZeroMemory transform, cbPtr
    End If

    CreateTranslateTransform_A = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks
    
End Function

Private Function CreateLinearAnimation(beginValue As Single, endValue As Single, beginTime As Single, endTime As Single, linearAnimation As IDCompositionAnimation) As Long
    PostLog "CreateLinearAnimation->Entry"
    On Error GoTo e0
    Dim hr As Long: hr = IIf(VarPtr(linearAnimation) = 0, E_POINTER, S_OK)
    
    If SUCCEEDED(hr) Then
        Set linearAnimation = Nothing
        hr = IIf(device Is Nothing, E_UNEXPECTED, S_OK)
    End If
    
    Dim animation As IDCompositionAnimation
    
    If SUCCEEDED(hr) Then
        device.CreateAnimation animation
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        If beginTime > 0 Then
            animation.AddCubic 0, beginValue, 0, 0, 0
            'hr = Err.LastHResult
        End If
    End If
    
    If SUCCEEDED(hr) Then
        animation.AddCubic beginTime, beginValue, (endValue - beginValue) / (endTime - beginTime), 0, 0
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        animation.End endTime, endValue
        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        PostLog "CreateLinearAnimation->animation.Detach()"
        CopyMemory linearAnimation, animation, cbPtr
        ZeroMemory animation, cbPtr
    End If
    PostLog "CreateLinearAnimation->hr=0x" & Hex$(hr)
    CreateLinearAnimation = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks

End Function





Private Function CreateSurface(ByVal size As Long, ByVal fRed As Single, ByVal fGreen As Single, ByVal fBlue As Single, surface As IDCompositionSurface) As Long
    PostLog "CreateSurface->Entry"
    On Error GoTo e0
    Dim hr As Long: hr = IIf(VarPtr(surface) = 0, E_POINTER, S_OK)
    
    If SUCCEEDED(hr) Then
        hr = IIf(((device Is Nothing) Or (d2d1Factory Is Nothing) Or (d2d1DeviceContext Is Nothing)), E_UNEXPECTED, S_OK)
        Set surface = Nothing
    End If
    
    Dim surfaceTile As IDCompositionSurface
    
    If SUCCEEDED(hr) Then
        device.CreateSurface size, size, DXGI_FORMAT_R8G8B8A8_UNORM, DXGI_ALPHA_MODE_IGNORE, surfaceTile
        'hr = Err.LastHResult
    End If
    
    Dim dxgiSurface As IDXGISurface
    Dim offset As POINT
    
    If SUCCEEDED(hr) Then
        PostLog "CreateSurface->BeginDraw"
        Dim RECT As RECT
        RECT.Right = size: RECT.Bottom = size
        
        surfaceTile.BeginDraw RECT, IID_IDXGISurface, dxgiSurface, offset
        'hr = Err.LastHResult
    End If
        
    Dim d2d1Bitmap As ID2D1Bitmap1
    If SUCCEEDED(hr) Then
        Dim dpiX As Single, dpiY As Single
        d2d1Factory.GetDesktopDpi dpiX, dpiY
        
        Dim bitmapProperties As D2D1_BITMAP_PROPERTIES1
        bitmapProperties.bitmapOptions = D2D1_BITMAP_OPTIONS_TARGET Or D2D1_BITMAP_OPTIONS_CANNOT_DRAW
        bitmapProperties.PixelFormat.Format = DXGI_FORMAT_R8G8B8A8_UNORM
        bitmapProperties.PixelFormat.AlphaMode = D2D1_ALPHA_MODE_IGNORE
        bitmapProperties.dpiX = dpiX
        bitmapProperties.dpiY = dpiY
        
        d2d1DeviceContext.CreateBitmapFromDxgiSurface dxgiSurface, bitmapProperties, d2d1Bitmap
        'hr = Err.LastHResult
        
        If SUCCEEDED(hr) Then
            d2d1DeviceContext.SetTarget d2d1Bitmap
        End If
        
        Dim d2d1Brush As ID2D1SolidColorBrush
        
        If SUCCEEDED(hr) Then
            Dim clr As D2D1_COLOR_F
            clr.r = fRed: clr.g = fGreen: clr.b = fBlue: clr.a = 1#
            Set d2d1Brush = d2d1DeviceContext.CreateSolidColorBrush(clr, ByVal 0)
            'hr = Err.LastHResult
        End If
        
        If SUCCEEDED(hr) Then
            PostLog "CreateSurface->DeviceContext::Draw"
            d2d1DeviceContext.BeginDraw
            
            Dim rf As D2D1_RECT_F
            rf.Left = offset.x
            rf.Top = offset.y
            rf.Right = offset.x + size
            rf.Bottom = offset.y + size
            
            d2d1DeviceContext.FillRectangle rf, d2d1Brush
            
            d2d1DeviceContext.EndDraw ByVal 0, ByVal 0
            'hr = Err.LastHResult
        End If
        PostLog "CreateSurface->EndDraw"
        surfaceTile.EndDraw
        
End If

If SUCCEEDED(hr) Then
    PostLog "CreateSurface->surfaceTile.Detach()"
    CopyMemory surface, surfaceTile, cbPtr
    ZeroMemory surfaceTile, cbPtr
End If

CreateSurface = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'We handle errors manually with the endless SUCCEEDED checks
   
End Function

Private Function UpdateVisuals(currentVisual As Long, nextVisual As Long) As LongPtr
    PostLog "UpdateVisuals(" & currentVisual & ", " & nextVisual & ")->Entry"
    On Error GoTo e0
    Dim hr As Long
    visualRight.SetContent surfaceLeftChild(nextVisual)
    ''hr = Err.LastHResult
    
    If SUCCEEDED(hr) Then
        effectGroupLeftChild(currentVisual).SetOpacity 1#
        ''hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        effectGroupLeftChild(nextVisual).SetOpacity 0#
        ''hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        PostLog "UpdateVisuals->Commit()"
        device.Commit
        ''hr = Err.LastHResult
    End If
    PostLog "UpdateVisuals->Out, hr=0x" & Hex$(hr)
    UpdateVisuals = IIf(SUCCEEDED(hr), 0, 1)
Exit Function
e0:
hr = Err.Number
Resume Next 'Handle errors manually

End Function

Private Function pvZoomOut() As Long
    PostLog "ZoomOut->Entry"
    On Error GoTo e0
    Dim hr As Long: hr = IIf(state = ZOOMEDOUT, E_UNEXPECTED, S_OK)
    
    If SUCCEEDED(hr) Then
        actionType = ACTION_TYPE.ZoomOut
        hr = SetEffectOnVisuals()
    End If
    
    If SUCCEEDED(hr) Then
        PostLog "ZoomOut->Commit()"
        device.Commit
'        'hr = Err.LastHResult
    End If
    
    If SUCCEEDED(hr) Then
        state = ZOOMEDOUT
    End If
    
    pvZoomOut = hr
Exit Function
e0:
hr = Err.Number
Resume Next 'Handle errors manually
End Function

Private Function pvZoomIn() As Long
    PostLog "ZoomIn->Entry"
    On Error GoTo e0
    Dim hr As Long: hr = IIf(state = ZOOMEDIN, E_UNEXPECTED, S_OK)
    
    If SUCCEEDED(hr) Then
        actionType = ACTION_TYPE.ZoomIn
        hr = SetEffectOnVisuals()
    End If
    
    If SUCCEEDED(hr) Then
        PostLog "ZoomIn->Commit()"
        device.Commit
    End If
    
    If SUCCEEDED(hr) Then
        state = ZOOMEDIN
    End If
    
    pvZoomIn = hr
    
Exit Function
e0:
hr = Err.Number
Resume Next 'Handle errors manually
End Function

Private Function OnLeftButton() As LongPtr
    Dim hr As Long
    If state = ZOOMEDOUT Then
        hr = pvZoomIn()
    Else
        hr = pvZoomOut()
    End If
    
    OnLeftButton = IIf(SUCCEEDED(hr), 0, 1)
End Function

Private Function OnKeyDown(wParam As LongPtr) As LongPtr
    PostLog "OnKeyDown(" & wParam & ")->Entry"
    Dim lr As LongPtr
    
    If state = ZOOMEDOUT Then
        If (wParam = vbKey1) And (currentVisual <> 0) Then
            lr = UpdateVisuals(currentVisual, 0)
            currentVisual = 0
        ElseIf (wParam = vbKey2) And (currentVisual <> 1) Then
            lr = UpdateVisuals(currentVisual, 1)
            currentVisual = 1
        ElseIf (wParam = vbKey3) And (currentVisual <> 2) Then
            lr = UpdateVisuals(currentVisual, 2)
            currentVisual = 2
        ElseIf (wParam = vbKey4) And (currentVisual <> 3) Then
            lr = UpdateVisuals(currentVisual, 3)
            currentVisual = 3
        End If
        PostLog "OnKeyDown->UpdateVisuals hr=0x" & Hex$(lr)
    End If
    
    OnKeyDown = lr
End Function


Private Function OnClose() As Long
    If m_hWnd Then
        DestroyWindow m_hWnd
        m_hWnd = 0
    End If
    
    OnClose = 0
End Function

Private Function OnDestroy() As Long
    PostQuitMessage 0
End Function

Private Function OnPaint() As Long
    If bpTrigger = False Then
        bpTrigger = True
        PostLog "OnPaint()"
    End If
    Dim rcClient As RECT
    Dim ps As PAINTSTRUCT
    Dim hDC As Long: hDC = BeginPaint(m_hWnd, ps)
    
    GetClientRect m_hWnd, rcClient
    
    Dim hLogo As Long: hLogo = CreateFontW(fontHeightLogo, 0, 0, 0, 0, CFalse, 0, 0, 0, 0, 0, 0, 0, StrPtr(FONT_TYPEFACE))
    If hLogo <> 0 Then
        Dim hOldFont As Long: hOldFont = SelectObject(hDC, hLogo)
   
        SetBkMode hDC, TRANSPARENT
        
        rcClient.Top = 10
        rcClient.Left = 50
        
        DrawTextW hDC, StrPtr("Windows samples"), -1, rcClient, DT_WORDBREAK
        
        SelectObject hDC, hOldFont
        
        DeleteObject hLogo
    End If
    
    Dim hTitle As LongPtr: hTitle = CreateFontW(fontHeightTitle, 0, 0, 0, 0, CFalse, 0, 0, 0, 0, 0, 0, 0, StrPtr(FONT_TYPEFACE))
    If hTitle <> 0 Then
        Dim hOldFontT As LongPtr: hOldFontT = SelectObject(hDC, hTitle)
        
        SetTextColor hDC, GetSysColor(COLOR_WINDOWTEXT)
        
        rcClient.Top = 25
        rcClient.Left = 50
        
        DrawTextW hDC, StrPtr("DirectComposition Effects Sample"), -1, rcClient, DT_WORDBREAK
        
        SelectObject hDC, hOldFontT
        
        DeleteObject hTitle
    End If
        
    Dim hDescription As LongPtr: hDescription = CreateFontW(fontHeightDescription, 0, 0, 0, 0, CFalse, 0, 0, 0, 0, 0, 0, 0, StrPtr(FONT_TYPEFACE)) '    // Description Font And Size
    If (hDescription <> 0) Then
        Dim hOldFontD As LongPtr: hOldFontD = SelectObject(hDC, hDescription)

        rcClient.Top = 90
        rcClient.Left = 50

        DrawTextW hDC, StrPtr("This sample explains how to use DirectComposition effects: rotation, scaling, perspective, translation and opacity."), -1, rcClient, DT_WORDBREAK

        rcClient.Top = 500
        rcClient.Left = 450

        DrawTextW hDC, StrPtr("A) Left-click to toggle between single and multiple-panels view." & vbCrLf & "B) Use keys 1-4 to switch the color of the right-panel."), -1, rcClient, DT_WORDBREAK

        SelectObject hDC, hOldFont

        DeleteObject hDescription
    End If
   
    Call EndPaint(m_hWnd, ps)
    
    OnPaint = 0
End Function
Private Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim result As Long
    
    Select Case uMsg
        Case WM_CREATE
            PostLog "WM_CREATE"
            
         Case WM_LBUTTONUP
            PostLog "WM_LBUTTONUP"
             result = OnLeftButton()
            
         Case WM_KEYDOWN
             result = OnKeyDown(wParam)
            
        Case WM_CLOSE
            result = OnClose()
            
        Case WM_DESTROY
            result = OnDestroy()
            
        Case WM_PAINT
            result = OnPaint()
            
        Case Else
            result = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End Select
    
    WindowProc = result
End Function

