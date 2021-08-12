VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Test DashBoard Controls"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11550
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   11520
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11550
      Begin VB.CommandButton Command5 
         Caption         =   "AxDashGraphLabel2"
         Height          =   360
         Left            =   7440
         TabIndex        =   4
         Top             =   60
         Width           =   1725
      End
      Begin VB.CommandButton Command4 
         Caption         =   "AxDashGraphLabel"
         Height          =   360
         Left            =   5631
         TabIndex        =   3
         Top             =   60
         Width           =   1725
      End
      Begin VB.CommandButton Command3 
         Caption         =   "AxDashBigLabel"
         Height          =   360
         Left            =   3824
         TabIndex        =   2
         Top             =   60
         Width           =   1725
      End
      Begin VB.CommandButton Command2 
         Caption         =   "AxDashSmallLabel"
         Height          =   360
         Left            =   2017
         TabIndex        =   1
         Top             =   60
         Width           =   1725
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
Form4.Show
End Sub

Private Sub Command5_Click()
Form5.Show
End Sub
