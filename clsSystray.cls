VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'System Tray Class

'''''''''''''''''''''''''
'the original code was by
'Christopher Lord
'heenix@project-angel.net
'17th March 2006
'''''''''''''''''''''''''
'
'_________________________________________________
'
'Debugged and enhanced by Ulli    umgedv@yahoo.com
'
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'______
'
'HOW TO
'
'¯¯¯¯¯¯
'
'   1 - add the class to your project
'   2 - declare an object variable with events for it
'   3 - instantiate it as New clsSystray
'
'   you now have the following functions (all return an error code or success)
'
'       Name                what it does
'       ----                --------------------------------
'
'       SetOwner            makes a form the owner of the tray-icon
'
'       AddIconToTray       adds an icon to the system tray (there can only be one tray-icon per class instance)
'
'       RemoveIconFromTray  removes the icon from the system tray
'
'       ShowBalloon         opens a balloon message attached to the tray-icon
'
'       HideBalloon         closes the current balloon message if one is showing
'
'
'   and the following properties (query LastErrorNumber to assert success)
'
'       Visible             property get - returns true if the icon has been sent to the tray
'
'       Enabled             property let - sets the icon to the Enabled state
'
'       Enabled             property get - returns true when the icon is enabled
'
'       Icon                property let - supplies a new icon to be displayed in tray
'
'       Tooltip             property let - supplies a new tooltip text for the tray-icon
'
'       Tooltip             property get - returns the current tooltip text
'
'       LastErrorNumber     property get - returns the last error number
'
'       LastErrorText       property get - returns a descriptive error text
'
'
'   events
'
'       MouseMove(Button As Integer)
'
'       DoubleClick(Button As Integer)
'
'       MouseDown(Button As Integer)
'
'       MouseUp(Button As Integer)
'
'       BalloonTimeout()
'
'       BalloonClicked()
'
'       BalloonShow()
'
'       BalloonHide()
'
'
'   Error Codes 'zero=success - positive=warning - negative=error
'
'       Success                  0
'
'       Not A Form              -1
'
'       No Parent Form          -2
'
'       No Icon Supplied        -3
'
'       No Suitable Menu        +4
'
'       Menu Not Hidden         +5
'
'       No Default Menu Item    +6
'
'       Icon Not Loaded         -7
'
'       Icon Disabled           -8
'
'       Icon Already Loaded     -9
'
'       No Legal Icon Id        -10 'refers to icon inside the balloon
'
'       No Balloon Text         -11
'
'       If the owner form has an invisible menu named mnuTray (code name; not caption) then
'       this will be displayed when you rightclick on the tray icon. A menu item under mnuTray
'       named mnuDefault (if that exists) will be displayed in bold font.
'       The MouseUp Event (vbRightButton) will be fired before showing the menu so that you have
'       a chance to modify the menu if that should be necessary.
'
Option Explicit

Private Declare Function NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconW" (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Long

Private Type NOTIFYICONDATAW
    icoSize                 As Long
    icoHwnd                 As Long
    icoId                   As Long
    icoFlags                As Long
    icoCallbackMessage      As Long
    icoSource               As Long
    icoTooltip(0 To 255)    As Byte
    icoState                As Long
    icoStateMask            As Long
    szInfo(0 To 511)        As Byte
    uTimeOutOrVersion       As Long
    szInfoTitle(0 To 127)   As Byte
    dwInfoFlags             As Long
End Type
Private IconData            As NOTIFYICONDATAW

Public Enum BalloonIconType
    BlankIcon = 0
    InfoIcon = 1
    WarningIcon = 2
    ErrorIcon = 3
    SoundOff = 16
End Enum
#If False Then
Private BlankIcon, InfoIcon, WarningIcon, ErrorIcon, SoundOff
#End If

Private Enum SystrayConsts
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_LBUTTONDBLCLK = &H203
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    WM_RBUTTONDBLCLK = &H206
    WM_USER = &H400
    NIN_BALLOONSHOW = WM_USER + 2
    NIN_BALLOONHIDE = WM_USER + 3
    NIN_BALLOONTIMEOUT = WM_USER + 4
    NIN_BALLOONUSERCLICK = WM_USER + 5
    NIF_MESSAGE = 1
    NIF_ICON = 2
    NIF_TIP = 4
    NIF_INFO = 16
    NIM_ADD = 0
    NIM_MODIFY = 1
    NIM_DELETE = 2
    NIM_SETVERSION = 4
End Enum
#If False Then
Private WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_RBUTTONDBLCLK, WM_USER, NIN_BALLOONSHOW, _
        NIN_BALLOONHIDE, NIN_BALLOONTIMEOUT, NIN_BALLOONUSERCLICK, NIF_ICON, NIF_MESSAGE, NIF_TIP, NIF_INFO, NIM_ADD, NIM_MODIFY, _
        NIM_DELETE, NIM_SETVERSION
#End If

Public Enum ErrorCode 'zero=success - positive=warning - negative=error
    Success = 0
    NotAForm = -1
    NoParentForm = -2
    NoIconSupplied = -3
    NoMenu = 4
    MenuNotHidden = 5
    NoDefaultMenuItem = 6
    IconNotLoaded = -7
    IconDisabled = -8
    IconAlreadyLoaded = -9
    NoLegalIconId = -10 'inside the balloon
    NoBalloonText = -11
End Enum
#If False Then
Private Success, NotAForm, NoParentForm, NoIconSupplied, NoMenu, MenuNotHidden, NoDefaultMenuItem, IconNotLoaded, IconDisabled, IconAlreadyLoaded, _
        NoLegalIconId, NoBalloonText
#End If ':)

'some text
Private Const DisabledText  As String = "The Tray-Icon is disabled."
Private Const NotLoadedText As String = "No Tray-Icon was loaded."

'my private variables
Private WithEvents myForm   As Form 'parent form
Attribute myForm.VB_VarHelpID = -1
Private myIsIconInTray      As Boolean 'there's an icon in the tray
Private myEnabled           As Boolean 'it is enabled (or not)
Private myTooltipText       As String 'the icons tooltiptext
Private myLastErrorNumber   As ErrorCode 'last error number

'events
Public Event MouseMove(Button As Integer)
Public Event DoubleClick(Button As Integer)
Public Event MouseDown(Button As Integer)
Public Event MouseUp(Button As Integer)
Public Event BalloonTimeout()
Public Event BalloonClicked()
Public Event BalloonShow()
Public Event BalloonHide()

Public Function AddIconToTray(ByVal hIcon As Long, Optional Tooltiptext As String = vbNullString, Optional Enable As Boolean = False) As ErrorCode

  Dim GoAhead   As Boolean

    GoAhead = True
    myLastErrorNumber = Success

    Select Case True

      Case myForm Is Nothing
        myLastErrorNumber = NoParentForm
        GoAhead = False

      Case hIcon = 0
        myLastErrorNumber = NoIconSupplied
        GoAhead = False

      Case myIsIconInTray
        myLastErrorNumber = IconAlreadyLoaded
        GoAhead = False

      Case myForm.mnuTray Is Nothing
        myLastErrorNumber = NoMenu

      Case myForm.mnuTray.Visible
        myLastErrorNumber = MenuNotHidden

      Case myForm.mnuDefault Is Nothing
        myLastErrorNumber = NoDefaultMenuItem

    End Select

    If GoAhead Then
        myIsIconInTray = True
        IconData.icoSource = hIcon
        NotifyIcon NIM_ADD, IconData
        IconData.uTimeOutOrVersion = 3
        NotifyIcon NIM_SETVERSION, IconData
        Tooltip = Tooltiptext
        myEnabled = Enable
    End If
    AddIconToTray = myLastErrorNumber

End Function

Private Sub Class_Initialize()

    With IconData
        .icoSize = Len(IconData)
        .icoCallbackMessage = WM_MOUSEMOVE
    End With 'IconData
    myLastErrorNumber = Success

End Sub

Private Sub Class_Terminate()

    If myIsIconInTray Then
        RemoveIconFromTray
    End If

End Sub

Public Property Let Enabled(nuEnabled As Boolean)

    If nuEnabled = False Then 'hide balloon if one should be showing
        myEnabled = True 'set to true just to make sure we can hide a balloon
        HideBalloon
    End If
    myEnabled = CBool(nuEnabled) 'and now to proper state
    myLastErrorNumber = Success

End Property

Public Property Get Enabled() As Boolean

    Enabled = myEnabled

End Property

Public Function HideBalloon() As ErrorCode

    If myIsIconInTray Then
        If myEnabled Then
            With IconData
                .icoFlags = NIF_INFO
                .szInfoTitle(0) = 0
                .szInfo(0) = 0
                .dwInfoFlags = BlankIcon
                RefreshIcon
                myLastErrorNumber = Success
            End With 'ICONDATA
          Else 'MYENABLED = FALSE/0
            myLastErrorNumber = IconDisabled
        End If
      Else 'myIsIconInTray = FALSE/0
        myLastErrorNumber = IconNotLoaded
    End If
    HideBalloon = myLastErrorNumber

End Function

Public Property Let Icon(hIcon As Long)

    IconData.icoSource = hIcon
    RefreshIcon
    myLastErrorNumber = Success

End Property

Public Property Get IsIconInTray() As Boolean

    IsIconInTray = myIsIconInTray

End Property

Public Property Get LastErrorNumber() As ErrorCode

    LastErrorNumber = myLastErrorNumber

End Property

Public Property Get LastErrorText() As String

    Select Case LastErrorNumber

      Case Success
        LastErrorText = "Success :-)"

      Case NotAForm
        LastErrorText = "This is not a Form."

      Case NoParentForm
        LastErrorText = "Try-Icon has no Parent Form."

      Case NoIconSupplied
        LastErrorText = "You must suppy an Icon."

      Case NoMenu
        LastErrorText = "No suitable Menu found."

      Case MenuNotHidden
        LastErrorText = "Menu is suitable but not hidden."

      Case NoDefaultMenuItem
        LastErrorText = "No default Menu Item found; bold font was not applied."

      Case IconNotLoaded
        LastErrorText = NotLoadedText

      Case IconDisabled
        LastErrorText = DisabledText

      Case IconAlreadyLoaded
        LastErrorText = "Tray-Icon was already loaded."

      Case NoLegalIconId
        LastErrorText = "Balloon Icon identifier illegal."

      Case NoBalloonText
        LastErrorText = "Balloon Text was missing."

    End Select

End Property

Private Sub myForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If myEnabled Then
        Select Case myForm.ScaleX(X, myForm.ScaleMode, vbPixels)

          Case WM_MOUSEMOVE
            RaiseEvent MouseMove(Button)

          Case WM_LBUTTONDBLCLK
            RaiseEvent DoubleClick(vbLeftButton)

          Case WM_RBUTTONDBLCLK
            RaiseEvent DoubleClick(vbRightButton)

          Case WM_LBUTTONDOWN
            RaiseEvent MouseDown(vbLeftButton)

          Case WM_RBUTTONDOWN
            RaiseEvent MouseDown(vbRightButton)

          Case WM_LBUTTONUP
            RaiseEvent MouseUp(vbLeftButton)

          Case WM_RBUTTONUP
            RaiseEvent MouseUp(vbRightButton)
            With myForm
                'no warnings here; we have to assume that the menu is okay
                If Not .mnuTray Is Nothing Then
                    If .mnuTray.Visible = False Then
                        If .mnuDefault Is Nothing Then
                            .PopupMenu .mnuTray
                          Else 'NOT .MNUDEFAULT...
                            .PopupMenu .mnuTray, DefaultMenu:=.mnuDefault 'has a default menu item
                        End If
                    End If
                End If
            End With 'MYFORM

          Case NIN_BALLOONSHOW
            RaiseEvent BalloonShow

          Case NIN_BALLOONHIDE
            RaiseEvent BalloonHide

          Case NIN_BALLOONTIMEOUT
            RaiseEvent BalloonTimeout

          Case NIN_BALLOONUSERCLICK
            RaiseEvent BalloonClicked

        End Select
    End If

End Sub

Private Sub RefreshIcon()

    IconData.icoFlags = IconData.icoFlags Or NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    NotifyIcon NIM_MODIFY, IconData

End Sub

Public Function RemoveIconFromTray() As ErrorCode

    If myIsIconInTray Then
        If myEnabled Then
            NotifyIcon NIM_DELETE, IconData
            myIsIconInTray = False
            myLastErrorNumber = Success
          Else 'MYENABLED = FALSE/0
            myLastErrorNumber = IconDisabled
        End If
      Else 'myIsIconInTray = FALSE/0
        myLastErrorNumber = IconNotLoaded
    End If
    RemoveIconFromTray = myLastErrorNumber

End Function

Public Function SetOwner(Frm As Form) As ErrorCode

    If TypeOf Frm Is Form Then
        Set myForm = Frm
        With IconData
            .icoHwnd = myForm.hWnd
            .icoSource = myForm.Icon.Handle
        End With 'ICONDATA
        myLastErrorNumber = Success
      Else 'NOT TYPEOF...
        myLastErrorNumber = NotAForm
    End If
    SetOwner = myLastErrorNumber

End Function

Public Function ShowBalloon(Message As String, Optional Title As String = vbNullString, Optional ByVal Style As BalloonIconType = BlankIcon, Optional ByVal TimeOut As Long = 2000) As ErrorCode

    If myIsIconInTray Then
        If myEnabled Then
            Select Case Style And Not SoundOff
              Case BlankIcon, InfoIcon, WarningIcon, ErrorIcon
                With IconData
                    .icoFlags = NIF_INFO
                    If Message = vbNullString Then
                        .szInfoTitle(0) = 0
                        .szInfo(0) = 0
                        IconData.dwInfoFlags = BlankIcon
                      Else 'NOT MESSAGE...
                        StringToArray Message, .szInfo
                        StringToArray Title, .szInfoTitle
                        .uTimeOutOrVersion = TimeOut 'may be modified by windows
                        .dwInfoFlags = Style
                    End If
                End With 'ICONDATA
                RefreshIcon
                myLastErrorNumber = Success
              Case Else
                myLastErrorNumber = NoLegalIconId
            End Select
          Else 'MYENABLED = FALSE/0
            myLastErrorNumber = IconDisabled
        End If
      Else 'myIsIconInTray = FALSE/0
        myLastErrorNumber = IconNotLoaded
    End If
    ShowBalloon = myLastErrorNumber

End Function

Private Sub StringToArray(sString As String, bArray() As Byte)

  Dim Bytes() As Byte
  Dim Pointer As Long

    If Len(sString) > 0 Then
        Bytes = sString
        For Pointer = 0 To UBound(Bytes)
            bArray(Pointer) = Bytes(Pointer)
            If Pointer = UBound(bArray) Then
                Exit For 'loop varying pointer
            End If
        Next Pointer
    End If
    bArray(Pointer) = 0 'terminating zero

End Sub

Public Property Get Tooltip() As String

    Tooltip = myTooltipText

End Property

Public Property Let Tooltip(Text As String)

    If Len(Trim$(Text)) Then
        myTooltipText = Text
        StringToArray Text, IconData.icoTooltip
        RefreshIcon
        myLastErrorNumber = Success
      Else 'LEN(TRIM$(TEXT)) = FALSE/0
        myLastErrorNumber = NoBalloonText
    End If

End Property

':) Ulli's VB Code Formatter V2.21.6 (2006-Aug-31 15:20)  Decl: 209  Code: 335  Total: 544 Lines
':) CommentOnly: 113 (20,8%)  Commented: 29 (5,3%)  Empty: 99 (18,2%)  Max Logic Depth: 7
