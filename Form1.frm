VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy Date Time"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ExtraTxt 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Text            =   "ExtraTxt"
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox OptionTxt 
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Text            =   "OptionTxt"
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add Extra Text"
      Height          =   1455
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Width           =   3255
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "Add Custom text Here....."
         Top             =   720
         Width           =   3015
      End
      Begin VB.CheckBox LastModifiedCheckBox 
         Caption         =   "Last Modified On"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Display Options"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
      Begin VB.OptionButton CSS 
         Caption         =   "CSS Comment"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton HTML 
         Caption         =   "HTML Comment"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Simple 
         Caption         =   "Simple"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      ToolTipText     =   "Copy to Clipboard"
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: CopyDateTime
'Project Version: 0.1.0
'Programming Platform: VB6 (Classic Visual Basic)
'Target OS: Microsoft Windows 32bit / 64bit (make sure VB6 runtime files are installed)
'Last Modified On: Sunday, 26.11.2023 - 11:13:16 AM PST
'Project Author: Usman Afzal (https://www.usmanafzal.pk); (e-Mail: ua7575@gmail.com)
'Project Description: A Simple VB6 program to Copy current Date and Time to Clipboard.

Option Explicit
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub LastModifiedCheckBox_Click()

'Label1.Caption = ""

If LastModifiedCheckBox.Value = 1 Then
    ExtraTxt.Text = "yes"
Else
    ExtraTxt.Text = "no"
End If

ConditionsToFollow

End Sub

Sub ConditionsToFollow()

If ExtraTxt.Text = "yes" And OptionTxt.Text = "simple" Then
Label1.Caption = "Last Modified On: " & Format$(Now, "dddd, dd.mm.yyyy") & " - " & Format$(Time, "hh:mm:ss AM/PM") & " PST"

ElseIf ExtraTxt.Text = "no" And OptionTxt.Text = "simple" Then
Label1.Caption = Format$(Now, "dddd, dd.mm.yyyy") & " - " & Format$(Time, "hh:mm:ss AM/PM") & " PST"

ElseIf ExtraTxt.Text = "yes" And OptionTxt.Text = "html" Then
Label1.Caption = "<!-- " & "Last Modified On: " & Format$(Now, "dddd, dd.mm.yyyy") & " - " & Format$(Time, "hh:mm:ss AM/PM") & " PST" & " -->"

ElseIf ExtraTxt.Text = "no" And OptionTxt.Text = "html" Then
Label1.Caption = "<!-- " & Format$(Now, "dddd, dd.mm.yyyy") & " - " & Format$(Time, "hh:mm:ss AM/PM") & " PST" & " -->"

ElseIf ExtraTxt.Text = "yes" And OptionTxt.Text = "css" Then
Label1.Caption = "/* " & "Last Modified On: " & Format$(Now, "dddd, dd.mm.yyyy") & " - " & Format$(Time, "hh:mm:ss AM/PM") & " PST" & " */"

ElseIf ExtraTxt.Text = "no" And OptionTxt.Text = "css" Then
Label1.Caption = "/* " & Format$(Now, "dddd, dd.mm.yyyy") & " - " & Format$(Time, "hh:mm:ss AM/PM") & " PST" & " */"
End If

End Sub

Private Sub Command1_Click()

ConditionsToFollow

CopyToClipBoard

End Sub

Private Sub CopyToClipBoard()
 
Dim retries As Integer
    
On Error GoTo Clip_Error

Clipboard.Clear

Clipboard.SetText (Label1.Caption)
    
Exit Sub
    
Clip_Error:
 
    If Err = 521 Then
        If retries > 10 Then
            MsgBox "Unable to access clipboard" & vbCrLf & "Try again later"
        Else
            retries = retries + 1
            Sleep 100
            Resume
        End If
    Else
        MsgBox Error$
    End If
        
End Sub

Private Sub Form_Activate()

Label1.Caption = Format$(Now, "dddd, dd.mm.yyyy") & " - " & Format$(Time, "hh:mm:ss AM/PM") & " PST"
OptionTxt.Text = "simple"
ExtraTxt.Text = "no"

End Sub

Private Sub Simple_Click()

Label1.Caption = Format$(Now, "dddd, dd.mm.yyyy") & " - " & Format$(Time, "hh:mm:ss AM/PM") & " PST"
OptionTxt.Text = "simple"

End Sub

Private Sub HTML_Click()

Label1.Caption = "<!-- " & Format$(Now, "dddd, dd.mm.yyyy") & " - " & Format$(Time, "hh:mm:ss AM/PM") & " PST" & " -->"
OptionTxt.Text = "html"

End Sub

Private Sub CSS_Click()

Label1.Caption = "/* " & Format$(Now, "dddd, dd.mm.yyyy") & " - " & Format$(Time, "hh:mm:ss AM/PM") & " PST" & " */"
OptionTxt.Text = "css"

End Sub
