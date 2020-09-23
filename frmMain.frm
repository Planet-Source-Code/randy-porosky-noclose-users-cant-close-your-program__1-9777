VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NoClose"
   ClientHeight    =   495
   ClientLeft      =   4710
   ClientTop       =   3555
   ClientWidth     =   1695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Close Program"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' Make this the only way to close this program
    End
End Sub

Private Sub Form_Load()
    ' The code may be used in any program you should
    ' desire. If you want, you can give me credits,
    ' but it's not a request or demand.
    
    ' Well, let's get started...
    
    ' First of all, we'll test if this instance is
    ' going to replace another one...
    If LCase$(Command$) = "replaceme" Then
        ' We are replacing another instance, so
        ' let's read the values from the Registration
        ' Database, that describes the placement of
        ' the window
        Me.Left = Val(GetSetting("NoClose", "Info", "Left", Str$(Me.Left)))
        Me.Top = Val(GetSetting("NoClose", "Info", "Top", Str$(Me.Top)))
        Me.Width = Val(GetSetting("NoClose", "Info", "Width", Str$(Me.Width)))
        Me.Height = Val(GetSetting("NoClose", "Info", "Height", Str$(Me.Height)))
    End If
    
    ' This makes sure that the form is correctly placed
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Okay, we need to figure out how the program is
    ' about to be closed.
    ' To perform this test, we will only test if the
    ' UnloadMode is different from 1, which means that
    ' we can still unload the form via project code.
    If UnloadMode <> 1 Then
        ' Now, this wasn't invoked via code, thus we
        ' must prevent the lockdown of the program
        ' and/or system.
        ' 1) Save the settings we want to keep, to
        '     the Registration Database
        SaveSetting "NoClose", "Info", "Left", Trim$(Str$(Me.Left))
        SaveSetting "NoClose", "Info", "Top", Trim$(Str$(Me.Top))
        SaveSetting "NoClose", "Info", "Width", Trim$(Str$(Me.Width))
        SaveSetting "NoClose", "Info", "Height", Trim$(Str$(Me.Height))
        ' 2) Start the other program, making sure, that
        '     the program will start even if it's placed
        '     in a root directory.
        If Right$(App.Path, 1) <> "\" Then
            Call Shell(App.Path & "\" & App.EXEName & ".exe replaceme", vbNormalFocus)
        Else
            Call Shell(App.Path & App.EXEName & ".exe replaceme", vbNormalFocus)
        End If
    End If
End Sub
