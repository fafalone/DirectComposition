# twinBASIC DirectComposition Demos

This repository is to show off some basic demos of using DirectComposition/Direct2D in twinBASIC. 

**Update (19 Dec 2023):** .twinproj updated to reference WinDevLib (formerly tbShellLib) 7.0-- this eliminates package errors that tB did not raise at the time this project was initially released.

### Requirements
-DirectX 11, which is a pre-installed standard on Windows 7 and newer.

-A relatively recent build of [twinBASIC](https://github.com/twinbasic/twinbasic) to build from source.

-If you create a new project based on this code, you'll need to add the tbShellLib package, version 5.0.203 or higher, which contains all the APIs and interfaces. It's available on the package server in the 'twinpack Packages'  list in your project settings -> COM Type Library / ActiveX References; [illustration here](https://github.com/fafalone/tbShellLib/issues/9). It's has already been added to the demo, so no action is neccessary for that, only for new projects.

## DirectComposition Effects Demo
The first demo is a just a basic proof of concept, a close-as-possible port of the [Microsoft DirectComposition Effects SDK example](https://github.com/microsoft/Windows-classic-samples/tree/main/Samples/DirectCompositionEffects). The app is provisioned for having more than one demo by running the Effects Sample in it's own thread, which a log sync'd back to the Launch Form inside a critical section.

![Screenshot](https://i.imgur.com/xr6jyOL.gif)

Since the demo is ported as close as possible, you'll find something additional of interest in this project: Instead of using a Form, it creates it's own window from scratch using API and handles the entire message pump (error handlers omitted):

```
    Private Function CreateApplicationWindow() As Long

        Dim hr As Long = S_OK
    
        Dim wcex As WNDCLASSEX
    
        wcex.cbSize = LenB(wcex)
        wcex.style = CS_HREDRAW Or CS_VREDRAW
        wcex.lpfnWndProc = AddressOf WindowProc
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

        If SUCCEEDED(hr) Then
            PostLog "RegisterClassEx succeeded"
            Dim RECT As RECT
            RECT.Right = windowWidth: RECT.Bottom = windowHeight
            AdjustWindowRect RECT, WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_MINIMIZEBOX, 0
        
            m_hWnd = CreateWindowExW(0, StrPtr(wndClass), StrPtr(wndName), _
                                WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_MINIMIZEBOX, _
                                CW_USEDEFAULT, CW_USEDEFAULT, RECT.Right - RECT.Left, RECT.Bottom - RECT.Top, _
                                0, 0, App.hInstance, ByVal 0)
        End If
```

```
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
            result = CLng(tMSG.wParam)
        End If
        
        EnterMessageLoop = result
    End Function
```

After that, we get into all the DirectComposition/Direct2D code, which is too complex to go into much detail here; but the basic steps are to start with the `D3D11CreateDevice` and `D2D1CreateFactory` APIs to create the root DirectX objects, get a DXGI interface from the former, then use the `DCompositionCreateDevice` to create the rendering object. After that, we create surfaces, make those into DirectComposition visuals, and apply various transform effects and animations. 

I recommend following the code starting from `BeforeEnteringMessageLoop` to see all the object creation, then following from OnKeyDown and OnLeftButton to see how it responds to the two actions.

Stay tuned for more demos!

