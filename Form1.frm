VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "Registration Form"
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   13245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsubmit 
      BackColor       =   &H0080FF80&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   3255
   End
   Begin VB.CommandButton cmdrefresh 
      BackColor       =   &H0080C0FF&
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   3255
   End
   Begin VB.TextBox Txtbirthday 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   6
      Top             =   3840
      Width           =   4335
   End
   Begin VB.TextBox Txtage 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   5
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox Txtname 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   4
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MADE BY: RAPHAEL ARNALDO CRUZ"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9360
      TabIndex        =   12
      Top             =   8040
      Width           =   3495
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   120
      X2              =   13080
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   120
      X2              =   13080
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   13080
      X2              =   13080
      Y1              =   120
      Y2              =   8520
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   8520
   End
   Begin VB.Label Lblbirthday 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   9
      Top             =   3840
      Width           =   4815
   End
   Begin VB.Label Lblage 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   8
      Top             =   2760
      Width           =   4815
   End
   Begin VB.Label Lblname 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   7
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROFILE REGISTRATION"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   10215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdrefresh_Click()
Txtname.Enabled = True
Txtage.Enabled = True
Txtbirthday.Enabled = True
cmdsubmit.Enabled = True
End Sub

Private Sub cmdsubmit_Click()
If Txtname = "" And Txtage = "" And Txtbirthday = "" Then
MsgBox "YOU NEED TO FILL UP THE NEEDED FIELDS", vbOKOnly = vbInformation, "Registration"
Else
If MsgBox("Are you sure do you want to submit your answer?", vbYesNo + vbQuestion, "Question #1") = vbYes Then
Lblname.Caption = Txtname.Text
Lblage.Caption = Txtage.Text
Lblbirthday.Caption = Txtbirthday.Text
Txtname.Enabled = False
Txtage.Enabled = False
Txtbirthday.Enabled = False
cmdsubmit.Enabled = False
End If
End If
End Sub

