VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000008&
   Caption         =   "Time"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTempo 
      Interval        =   1000
      Left            =   3855
      Top             =   2115
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   4
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1320
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   495
      Width           =   6900
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OldCaption As String

Private Sub tmrTempo_Timer()
   Dim msg As String
   msg = Time$
   If msg <> Caption Then
      txtTime.Text = msg
   End If
End Sub
