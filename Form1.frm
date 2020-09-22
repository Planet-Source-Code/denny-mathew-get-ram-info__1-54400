VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RAM Info"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&About"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":08CA
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Available Physical Memory: "
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Physical Memory: "
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type MEMORYSTATUS
   dwLength As Long
   dwMemoryLoad As Long
   dwTotalPhys As Long
   dwAvailPhys As Long
   dwTotalPageFile As Long
   dwAvailPageFile As Long
   dwTotalVirtual As Long
   dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As _
   MEMORYSTATUS)
Dim memInfo As MEMORYSTATUS



Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
MsgBox "RAMInfo v1.0" & vbCrLf & vbCrLf & "Copyright © 1998-2004 Denny Mathew.", vbInformation, "About"
End Sub

Private Sub Timer1_Timer()
Dim memsts As MEMORYSTATUS
Dim memory&
Dim msg$

GlobalMemoryStatus memsts
memory& = memsts.dwTotalPhys
Label4 = Format$(memory& \ 1024, "###,###,###") + "KB"
memory& = memsts.dwAvailPhys
Label5 = Format$(memory& \ 1024, "###,###,###") + "KB"
Timer1.Enabled = True
End Sub
