VERSION 5.00
Begin VB.Form FormTexture 
   Caption         =   "Texture creator"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Texture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7380
      Left            =   0
      Picture         =   "FormTexture.frx":0000
      ScaleHeight     =   7320
      ScaleWidth      =   7320
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   7380
   End
   Begin VB.PictureBox BarraS 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Left            =   10800
      ScaleHeight     =   3075
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox BarraD 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Left            =   11160
      ScaleHeight     =   3075
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Front 
      Height          =   3135
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image Back 
      Height          =   3135
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "FormTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
