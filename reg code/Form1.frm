VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration Validation"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   2085
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   4320
      Width           =   5145
   End
   Begin VB.Frame Frame2 
      Caption         =   "Validate Key"
      Height          =   1725
      Left            =   90
      TabIndex        =   6
      Top             =   1980
      Width           =   5145
      Begin VB.CommandButton Command2 
         Caption         =   "&Validate"
         Height          =   285
         Left            =   3060
         TabIndex        =   11
         Top             =   1260
         Width           =   1725
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1350
         TabIndex        =   10
         Top             =   810
         Width           =   3435
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1350
         TabIndex        =   8
         Top             =   360
         Width           =   3435
      End
      Begin VB.Label Label4 
         Caption         =   "Key:"
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   810
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Username:"
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generate Key"
      Height          =   1725
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5145
      Begin VB.CommandButton Command1 
         Caption         =   "&Generate"
         Height          =   285
         Left            =   3060
         TabIndex        =   5
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1350
         TabIndex        =   4
         Top             =   810
         Width           =   3525
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1350
         TabIndex        =   2
         Top             =   360
         Width           =   3525
      End
      Begin VB.Label Label2 
         Caption         =   "Key:"
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   810
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   90
      TabIndex        =   12
      Top             =   3870
      Width           =   5145
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GenerateKey(InputString As TextBox, OutputString As TextBox)
Dim Al01 As String, sAl01 As String, tAl01 As String
Dim Al02 As String, sAl02 As String, tAl02 As String
Dim Al03 As String, sAl03 As String, tAl03 As String
Dim AlFinal As String

Al01 = Left(InputString, 1)
sAl01 = Asc(Al01)
tAl01 = sAl01 * 2

Al02 = Mid(InputString, 2, 1)
sAl02 = Asc(Al02)
tAl02 = sAl02 * 4

Al03 = Right(InputString, 1)
sAl03 = Asc(Al03)
tAl03 = sAl03 * 3

AlFinal = tAl01 & "-" & tAl02 & "-" & tAl03

OutputString = AlFinal
End Function


Private Function ValidateKey(InputKey As TextBox, UserString As TextBox)
Dim Al01 As String, sAl01 As String, tAl01 As String
Dim Al02 As String, sAl02 As String, tAl02 As String
Dim Al03 As String, sAl03 As String, tAl03 As String
Dim RetS1 As String, RetS2 As String, RetS3 As String
Dim Step1 As Boolean, Step2 As Boolean, Step3 As Boolean

Al01 = Left(InputKey, 3)
Text5.Text = Text5.Text & "Key returns:" & vbCrLf & "Al01 = " & Al01
sAl01 = Al01 / 2
Text5.Text = Text5.Text & vbCrLf & "sAl01 = " & sAl01
tAl01 = Chr(sAl01)
Text5.Text = Text5.Text & vbCrLf & "tAl01 = " & tAl01
RetS1 = Left(UserString, 1)
Text5.Text = Text5.Text & vbCrLf & "RetS1 = " & RetS1
    If tAl01 = RetS1 Then
    Step1 = True
    Text5.Text = Text5.Text & vbCrLf & "Step1 = true"
    Else
    Step1 = False
    Text5.Text = Text5.Text & vbCrLf & "Step1 = false"
    End If
    
If Step1 = True Then
    Al02 = Mid(InputKey, 5, 3)
    Text5.Text = Text5.Text & vbCrLf & "Al02 = " & Al02
    sAl02 = Al02 / 4
    Text5.Text = Text5.Text & vbCrLf & "sAl02 = " & sAl02
    tAl02 = Chr(sAl02)
    Text5.Text = Text5.Text & vbCrLf & "tAl02 = " & tAl02
    RetS2 = Mid(UserString, 2, 1)
    Text5.Text = Text5.Text & vbCrLf & "RetS2 = " & RetS2
        If tAl02 = RetS2 Then
        Step2 = True
        Text5.Text = Text5.Text & vbCrLf & "Step2 = true"
        Else
        Step2 = False
        Text5.Text = Text5.Text & vbCrLf & "Step2 = false"
        End If
End If

    If Step2 = True Then
        Al03 = Right(InputKey, 3)
        Text5.Text = Text5.Text & vbCrLf & "Al03 = " & Al03
        sAl03 = Al03 / 3
        Text5.Text = Text5.Text & vbCrLf & "sAl03 = " & sAl03
        tAl03 = Chr(sAl03)
        Text5.Text = Text5.Text & vbCrLf & "tAl03 = " & tAl03
        RetS3 = Right(UserString, 1)
        Text5.Text = Text5.Text & vbCrLf & "RetS3 = " & RetS3
            If tAl03 = RetS3 Then
            Step3 = True
            Text5.Text = Text5.Text & vbCrLf & "Step3 = true"
            Else
            Step3 = False
            Text5.Text = Text5.Text & vbCrLf & "Step3 = false"
            End If
    End If
    
        If Step3 = True Then
        Label5.Caption = "Valid"
        Text5.Text = Text5.Text & vbCrLf & "Key returned - valid."
        Else
        Label5.Caption = "Invalid"
        Text5.Text = Text5.Text & vbCrLf & "Key returned - invalid."
        End If
        
            
    


End Function



Private Sub Command1_Click()
GenerateKey Text1, Text2
End Sub

Private Sub Command2_Click()
Label5.Caption = ""
Text5.Text = ""
ValidateKey Text4, Text3
End Sub

