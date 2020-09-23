VERSION 5.00
Begin VB.Form frmSqrReduce 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sqr Reduce"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdReduce 
      Caption         =   "&Reduce"
      Default         =   -1  'True
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   ")"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sqr("
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   285
   End
End
Attribute VB_Name = "frmSqrReduce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrimeArray&(4791)
Private Function IsOdd(ByVal lNum&) As Boolean
    IsOdd = -(lNum And 1)
End Function
Private Sub cmdclear_click()
    txtNumber = ""
    txtNumber.SetFocus
End Sub
Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdReduce_Click()
    If txtNumber = "" Then Exit Sub
    If Not IsNumeric(txtNumber) Then
        txtNumber.SetFocus
        txtNumber.SelStart = 0
        txtNumber.SelLength = Len(txtNumber)
        Exit Sub
    End If
    On Error GoTo Hell
    Dim Num&
    Num = Val(txtNumber)
    If Num Mod CLng(Sqr(Num)) = 0 And InStr(CStr(Sqr(Num)), ".") = 0 Then
        MsgBox "Sqr(" & txtNumber & ") = " & CStr(Sqr(Num))
        cmdClear.SetFocus
        Exit Sub
    End If
    Dim NumPrimes%(), X&
    ReDim NumPrimes%(Num)
    Do
        If Num Mod PrimeArray(X) = 0 Then
            NumPrimes(X) = NumPrimes(X) + 1
            Num = Num / PrimeArray(X)
            X = X - 1
        End If
        X = X + 1
        If X > Num Then X = 0
    Loop Until Num = 1
    Dim UnderSqr&, OutOfSqr&
    UnderSqr = 1
    OutOfSqr = 1
    X = 0
    Do
        If NumPrimes(X) <> 0 Then
            If Not IsOdd(NumPrimes(X)) Then
                OutOfSqr = OutOfSqr * Sqr(PrimeArray(X) ^ NumPrimes(X))
            ElseIf IsOdd(NumPrimes(X)) And NumPrimes(X) <> 1 Then
                OutOfSqr = OutOfSqr * Sqr(PrimeArray(X) ^ (NumPrimes(X) - 1))
                UnderSqr = UnderSqr * PrimeArray(X)
            Else
                UnderSqr = UnderSqr * PrimeArray(X)
            End If
        End If
        X = X + 1
    Loop Until ((OutOfSqr ^ 2) * UnderSqr) = Val(txtNumber)
    Dim Message$
    Message = "Sqr(" & txtNumber & ") = "
    If OutOfSqr = 1 Then
        Message = Message & "Sqr(" & CStr(UnderSqr) & ")"
    ElseIf UnderSqr = 1 Then
        Message = Message & CStr(OutOfSqr)
    Else
        Message = Message & CStr(OutOfSqr) & "*Sqr(" & CStr(UnderSqr) & ")"
    End If
    MsgBox Message
    cmdClear.SetFocus
    Exit Sub
Hell:
    If Err.Number = 6 Or Err.Number = 7 Then MsgBox "You typed in a number too large for this program to work with.  Please input a smaller number.", vbInformation, "Smaller Num Please" Else MsgBox "Unexpected error " & CStr(Err.Number) & vbNewLine & Err.Description, vbCritical, "Error " & CStr(Err.Number)
    cmdclear_click
    Exit Sub
End Sub
Private Sub Form_Load()
    Dim FF%, PrimeLongFileName$, X%
    FF = FreeFile
    PrimeLongFileName = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & "primenumslong.txt"
    Open PrimeLongFileName For Random As #FF
    For X = 0 To 4791
        Get #FF, (X + 1), PrimeArray(X)
    Next X
End Sub
Private Sub txtNumber_KeyPress(KeyAscii%)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
Private Sub txtNumber_LostFocus()
    If txtNumber = "" Then Exit Sub
    If Not IsNumeric(txtNumber) Then
        txtNumber.SetFocus
        txtNumber.SelStart = 0
        txtNumber.SelLength = Len(txtNumber)
    End If
End Sub
