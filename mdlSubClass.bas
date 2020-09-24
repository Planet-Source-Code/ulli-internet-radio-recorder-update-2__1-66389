Attribute VB_Name = "modlSubClass"
Option Explicit

Private Declare Function SetNewHook Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallPrvHook Lib "user32" Alias "CallWindowProcA" (ByVal Prev As Long, ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const SWL_WNDPROC   As Long = -4
Private Type tHandleAndHook
    PrevHook                As Long
    SavdhWnd                As Long
End Type
Private HandlesAndHooks()   As tHandleAndHook
Private Idx                 As Long

Private HookedMsg           As Long

Public Sub Hook(Message As Long, hWnds As Variant) 'hWnds is an array of hWnds

    HookedMsg = Message
    ReDim HandlesAndHooks(0 To UBound(hWnds))
    For Idx = 0 To UBound(hWnds)
        With HandlesAndHooks(Idx)
            .SavdhWnd = hWnds(Idx)
            .PrevHook = SetNewHook(.SavdhWnd, SWL_WNDPROC, AddressOf MsgRelay)
        End With 'HANDLESANDHOOKS(Idx)
    Next Idx

End Sub

Private Function MsgRelay(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  'forward all messages except hooked messages which are consumed

    If nMsg <> HookedMsg Then
        For Idx = 0 To UBound(HandlesAndHooks)
            If hWnd = HandlesAndHooks(Idx).SavdhWnd Then
                MsgRelay = CallPrvHook(HandlesAndHooks(Idx).PrevHook, hWnd, nMsg, wParam, lParam)
                Exit For 'loop varying idx
            End If
        Next Idx
    End If

End Function

Public Sub UnHook()

    For Idx = UBound(HandlesAndHooks) To 0 Step -1
        With HandlesAndHooks(Idx)
            If .PrevHook Then
                SetNewHook .SavdhWnd, SWL_WNDPROC, .PrevHook
            End If
        End With 'HANDLESANDHOOKS(Idx)
    Next Idx
    ReDim HandlesAndHooks(0)

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Aug-31 15:20)  Decl: 13  Code: 44  Total: 57 Lines
':) CommentOnly: 1 (1,8%)  Commented: 3 (5,3%)  Empty: 12 (21,1%)  Max Logic Depth: 4
