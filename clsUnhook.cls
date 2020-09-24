VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnhook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_Terminate()

    UnHook

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Aug-31 15:20)  Decl: 1  Code: 9  Total: 10 Lines
':) CommentOnly: 0 (0%)  Commented: 0 (0%)  Empty: 3 (30%)  Max Logic Depth: 1
