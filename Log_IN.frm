VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form Log_IN 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   FillColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerNewACD 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   9480
   End
   Begin VB.Frame MessageFrame 
      BorderStyle     =   0  'None
      Caption         =   "&H00800000&"
      Height          =   2655
      Left            =   2400
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   10815
      Begin VB.Label TestoMessaggio 
         BackStyle       =   0  'Transparent
         Caption         =   $"Log_IN.frx":0000
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   480
         TabIndex        =   62
         Top             =   720
         Width           =   9615
      End
      Begin VB.Label TitoloMessaggio 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LOG-IN"
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   10455
      End
      Begin VB.Shape Shape4 
         Height          =   255
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "JukeBox© 1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   5
         FillColor       =   &H00B65610&
         FillStyle       =   0  'Solid
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   10455
      End
      Begin VB.Shape Shape18 
         FillStyle       =   0  'Solid
         Height          =   2055
         Left            =   600
         Top             =   600
         Width           =   10335
      End
      Begin VB.Shape ERCorrection 
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   2
         Left            =   -240
         Top             =   2280
         Width           =   855
      End
      Begin VB.Shape ERCorrection 
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   3
         Left            =   10200
         Top             =   -120
         Width           =   735
      End
   End
   Begin VB.Frame ErrorMessage 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   2400
      TabIndex        =   58
      Top             =   4320
      Visible         =   0   'False
      Width           =   10815
      Begin VB.TextBox ERChiudi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   450
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Chiudi"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Shape Shape20 
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   8640
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Shape ERCorrection 
         BorderColor     =   &H00800000&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   10
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   1
         Left            =   10560
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape ERCorrection 
         BorderColor     =   &H00800000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   0
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label ERTesto 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1095
         Left            =   600
         TabIndex        =   60
         Top             =   600
         Width           =   8775
      End
      Begin VB.Label ERTitolo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1560
         TabIndex        =   59
         Top             =   120
         Width           =   7575
      End
      Begin VB.Shape Shape16 
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   2415
         Left            =   0
         Top             =   0
         Width           =   10575
      End
      Begin VB.Shape Shape17 
         FillStyle       =   0  'Solid
         Height          =   2175
         Left            =   480
         Top             =   480
         Width           =   10335
      End
   End
   Begin VB.Frame FrameNome 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7455
      Left            =   3000
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Frame FramePNU 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6735
         Left            =   0
         TabIndex        =   50
         Top             =   840
         Visible         =   0   'False
         Width           =   9735
         Begin VB.TextBox CNP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00400000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "ChickenScratch"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   3
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   57
            Text            =   "Cancella"
            Top             =   4920
            Width           =   2295
         End
         Begin VB.TextBox CNP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00400000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "ChickenScratch"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   4
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "Conferma"
            Top             =   4920
            Width           =   2295
         End
         Begin VB.TextBox CNP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00B65610&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   4080
            PasswordChar    =   "*"
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   3240
            Width           =   5175
         End
         Begin VB.TextBox CNP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00B65610&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   4080
            PasswordChar    =   "*"
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1920
            Width           =   5175
         End
         Begin VB.Shape Shape15 
            FillStyle       =   0  'Solid
            Height          =   615
            Left            =   2400
            Top             =   5040
            Width           =   2295
         End
         Begin VB.Shape Shape14 
            FillStyle       =   0  'Solid
            Height          =   615
            Left            =   5280
            Top             =   5040
            Width           =   2295
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Conferma Password :"
            BeginProperty Font 
               Name            =   "ChickenScratch"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   975
            Left            =   120
            TabIndex        =   55
            Top             =   3240
            Width           =   3855
         End
         Begin VB.Shape Shape13 
            FillStyle       =   0  'Solid
            Height          =   615
            Left            =   4560
            Top             =   3480
            Width           =   4935
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Password :"
            BeginProperty Font 
               Name            =   "ChickenScratch"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   615
            Left            =   840
            TabIndex        =   53
            Top             =   2040
            Width           =   3735
         End
         Begin VB.Shape Shape12 
            FillStyle       =   0  'Solid
            Height          =   615
            Left            =   4440
            Top             =   2160
            Width           =   4935
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   $"Log_IN.frx":00B3
            BeginProperty Font 
               Name            =   "ChickenScratch"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D3C64E&
            Height          =   1095
            Left            =   480
            TabIndex        =   51
            Top             =   240
            Width           =   8775
         End
         Begin VB.Shape Shape11 
            BorderWidth     =   5
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   6735
            Left            =   0
            Top             =   -120
            Width           =   9735
         End
      End
      Begin VB.TextBox NuovoNome 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00B65610&
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   4080
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   6480
         Width           =   5175
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   29
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "FINE"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   28
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "minusc"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   27
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "back"
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   26
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "space"
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   25
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Z"
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   24
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Y"
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   23
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "X"
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   22
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "W"
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   21
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "V"
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   20
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "U"
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   19
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "T"
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   18
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "S"
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   17
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "R"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   16
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "Q"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   15
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "P"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   14
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "O"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   13
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "N"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   12
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "M"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   11
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "L"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   10
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "K"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   9
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "J"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   8
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "I"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   7
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "H"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   6
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "G"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   5
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "F"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   4
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "E"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   3
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "D"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   2
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "C"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   1
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "B"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Lettera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "A"
         Top             =   2280
         Width           =   855
      End
      Begin VB.Shape Shape10 
         Height          =   255
         Left            =   120
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "JukeBox© 1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"Log_IN.frx":01AD
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D3C64E&
         Height          =   1095
         Left            =   480
         TabIndex        =   48
         Top             =   840
         Width           =   8775
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   4440
         Top             =   6720
         Width           =   4935
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H000000FF&
         FillColor       =   &H00400000&
         FillStyle       =   0  'Solid
         Height          =   4095
         Left            =   960
         Top             =   2040
         Width           =   7815
      End
      Begin VB.Shape Shape8 
         FillStyle       =   0  'Solid
         Height          =   3735
         Left            =   1560
         Top             =   2640
         Width           =   7455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Utente :"
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   46
         Top             =   6600
         Width           =   3735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ">>> Creazione nuovo utente <<<"
         BeginProperty Font 
            Name            =   "ChickenScratch"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   15
         Top             =   120
         Width           =   7575
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   3
         FillColor       =   &H00B65610&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   9735
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   5
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   7455
         Left            =   0
         Top             =   120
         Width           =   9735
      End
   End
   Begin VB.PictureBox CDFermo 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   9840
      Picture         =   "Log_IN.frx":0271
      ScaleHeight     =   1335
      ScaleWidth      =   1335
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Password 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   5040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   6960
      Width           =   5535
   End
   Begin VB.FileListBox ListaFile 
      Height          =   1065
      Left            =   120
      TabIndex        =   2
      Top             =   10320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer CheckDisc 
      Interval        =   100
      Left            =   120
      Top             =   10440
   End
   Begin VB.ComboBox DrivesDisponibili 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS VoceLogIN 
      Height          =   375
      Left            =   600
      OleObjectBlob   =   "Log_IN.frx":2DA1
      TabIndex        =   0
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Benvenuto..."
      BeginProperty Font 
         Name            =   "ChickenScratch"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   3
      Left            =   5040
      TabIndex        =   13
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label NomeUtente 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ChickenScratch"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4680
      TabIndex        =   12
      Top             =   4680
      Width           =   6255
   End
   Begin VB.Label UtenteValido 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ChickenScratch"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   3480
      Width           =   3855
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash CDMoto 
      Height          =   1335
      Left            =   9840
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
      _cx             =   2566
      _cy             =   2355
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Transparent"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin VB.Shape Shape4 
      Height          =   255
      Index           =   1
      Left            =   4200
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "JukeBox© 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Descrizione 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ChickenScratch"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   4920
      TabIndex        =   4
      Top             =   5640
      Width           =   6015
   End
   Begin VB.Shape Finestra 
      BorderWidth     =   2
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   6135
      Left            =   4080
      Top             =   2520
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1575
      Index           =   1
      Left            =   0
      Top             =   9960
      Width           =   15375
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1575
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   15375
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   5895
      Left            =   4920
      Top             =   3120
      Width           =   6855
   End
End
Attribute VB_Name = "Log_IN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dichiarazione di due funzioni API che permetteranno di mostrare o nascondere il
'cursore all'interno di una o più textbox
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
'Dichiarazione dell'oggetto nero che permetterà al programma di masterizzare CD
'e ACESS CD
Public WithEvents Nero As Nero
Attribute Nero.VB_VarHelpID = -1
'Dichiaro l'oggetto che servirà al programma per capire quali masterizzatori si
'hanno a disposizione per la masterizzazione
Public Drives As INeroDrives
Attribute Drives.VB_VarHelpID = -1
'Dichiaro una variabile che servirà al programma per capire su quale masterizzatore
'dovrà andare a scrivere
Public WithEvents SelectedDrive As NeroDrive
Attribute SelectedDrive.VB_VarHelpID = -1
'Dichiaro un oggetto che servirà al programma per masterizzare quanto necessario
'sul Jukebox AccessCD
Public Traccia As New NeroISOTrack
Attribute Traccia.VB_VarHelpID = -1
'Dichiaro una variabile che servirà al programma per creare la directory root all'interno del cd
Public Cartella As New NeroFolder
'Dichiaro una variabile che servirà al programma per identificare il file contenente
'il numero ID del nuovo utente, nella funzione di creazione nuovo jukebox AccessCD
Public FileID As New NeroFile
'Dichiaro una variabile che servirà a contenere il valore della password
Dim ValPassword As String
'Dichiaro una variabile che servirà a salvare il nuovo nome definito
Dim ValNuovoNome As String
'Dichiaro una variabile che servirà a risolvere un BUG con la tastiera
Dim BUG As Boolean

Private Sub CNP_Change(Index As Integer)
    'Dichiarazione di una variabile temporanea
    Dim car As String
    'Dichiarazione di una seconda variabile temporanea
    Dim tmp As String
    'Se si ha premuto il tasto 1 o 3 (REALI) del pad numerico
    If BUG = True Then
        'Chiamo la funzione addetta alla correzione del campo
        Correggi_Campo CNP(Index)
    End If
End Sub

Private Sub CNP_GotFocus(Index As Integer)
    Dim tmp As String
    Dim I As Integer
    'Se la textbox selezionata è utilizzata come pulsante, allora...
    If Index > 2 Then
        'Nascondo il cursore della textbox selezionata
        HideCaret CNP(Index).hwnd
    End If
    'Imposto il colore di scrittura a giallo
    CNP(Index).ForeColor = vbYellow
    'Posiziono il cursore della taextbox alla fine della textbox
    CNP(Index).SelStart = Len(CNP(Index))
End Sub

Private Sub CNP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Dichiaro una variabile temporanea
    Dim tmp As String
    Select Case KeyCode
    'Se il tasto premuto è lo 0, allora...
    Case 96:
        'Viene verificato su quale controllo è stato premuto
        Select Case Index
        'Se è stato premuto sul tasto Cancella,allora...
        Case 3:
            'Cancello il contenuo delle due Textbox
            CNP(1) = ""
            CNP(2) = ""
        'Se è stato premuto il pulsante di conferma, allora...
        Case 4:
            'Richiamo la funzione addetta alla verifica della nuova password
            verifica_Nuova_Password
        End Select
    Case 97:
        'Se la textbox selezionata è un campo password, allora
        If Index <= 2 Then
            BUG = True
        End If
        On Error Resume Next
        CNP(Index - 1).SetFocus
        If Err.Number <> 0 Then
            CNP(4).SetFocus
        End If
    Case 99:
        'Se la textbox selezionata è un campo password, allora
        If Index <= 2 Then
            BUG = True
        End If
        On Error Resume Next
        CNP(Index + 1).SetFocus
        If Err.Number <> 0 Then
            CNP(1).SetFocus
        End If
    End Select
End Sub

Private Sub CNP_LostFocus(Index As Integer)
    'Imposto il colore di scrittura a bianco
    CNP(Index).ForeColor = vbWhite
End Sub

Private Sub ERChiudi_GotFocus()
    'Nascondo il cursore
    HideCaret False
End Sub

Private Sub ERChiudi_KeyDown(KeyCode As Integer, Shift As Integer)
    'Se è stato premuto il pulsante OK, allora...
    If KeyCode = 96 Then
        'Chiudo il messaggio di errore
        ErrorMessage.Visible = False
    End If
End Sub

Private Sub Form_Load()
    'Nascondo il cursore del mouse
    'ShowCursor False
    'Setto il colore di sfondo del form
    Me.BackColor = RGB(0, 0, 66)
    'Imposto il colore delle correzioni del messaggio di errore...
    ERCorrection(0).FillColor = Me.BackColor
    ERCorrection(1).FillColor = Me.BackColor
    '...e del messaggio normale
    ERCorrection(2).FillColor = Me.BackColor
    ERCorrection(3).FillColor = Me.BackColor
    'Imposto il colore di sfondo della finestra
    Finestra.FillColor = RGB(0, 0, 128)
    CDMoto.Movie = App.Path & "\FLASH\CD.swf"
    'Setto la voce del computer
    VoceLogIN.CurrentMode = 4
    'Setto l'indice dell'array contenente i numero ID dei cd preferiti
    IPref = 1
    'Viene richiamata la funzione addetta alla ricerca dei drives disponibili
    Ricava_Drives
    If Aperto = False Then
        'Mostro il messaggio di Log-IN
        Messaggio "LOG-IN", "Prego, inserisca il suo jukebox AccessCD all'interno dell'apposito vano CD, al fine di permettere il suo riconoscimento e consentirle così, di accedere alla lista dei suoi CD."
        'Esegui il messaggio di benvenuto
        VoceLogIN.Speak "Welcome!            Please insert your access CD!"
        'Viene aperto il drive per consentire l'inserimento del Jukebox Access CD
        SelectedDrive.EjectCD
    End If
    CheckDisc.Enabled = True
End Sub

Private Sub Ricava_Drives()
    'Dichiarazione di un indice temporaneo
    Dim I As Integer
    'Setto l'oggetto Nero
    Set Nero = New Nero
    'Setto la variabile Drives
    Set Drives = Nero.GetDrives(NERO_MEDIA_CDR)
    'Ciclo di verifica drives
    For I = 0 To Drives.Count - 1
        DrivesDisponibili.AddItem Drives(I).DeviceName & "     " & Drives(I).DriveLetter
    'Si passa ad analizzare il drive successivo
    Next
    '------------------------------------------------------------
    ' Setto il masterizzatore selezionato, il primo trovato
    '------------------------------------------------------------
    DrivesDisponibili.ListIndex = 0
    Set SelectedDrive = Drives(DrivesDisponibili.ListIndex)
End Sub

Private Sub CheckDisc_Timer()
    'Viene verificato se il cd è stato inserito e se così fosse, setta la path
    'dell'oggetto ListaFile alla radice del cd
    On Error Resume Next
    ListaFile.Path = SelectedDrive.DriveLetter & ":\"
    'Se non si è verificato nessun errore, vuol dire che il cd è stato inserito,
    'quindi...
    If Err.Number = 0 Then
        'Chiamo la funzione addetta al controllo del cd inserito
        Verifica_CD
        'Nascondo l'immagine del cd fermo
        CDFermo.Visible = False
        'Nascondo il frame contenente il messaggio di InsertCD
        MessageFrame.Visible = False
        'Fermo il timer
        CheckDisc.Enabled = False
        'Esegui il messaggio di inserisci password
        VoceLogIN.Speak "Please insert your password!!!"
        'Setto il focus al campo Password
        Password.SetFocus
    End If
End Sub

Private Sub Verifica_CD()
    'Dichiarazione di un indice temporaneo
    Dim I As Integer
    'Dichiaro una variabile che servirà a contenere le righe estratte dai file
    Dim Riga As String
    'Viene dichiarata una variabile che servirà al programma per capire se il file
    'contenente il numero ID dell'utente è stato trovato
    Dim UtenteTrovato As Boolean
    'Dichiaro una variabile che servirà al programma per capire se il numeroID trovato
    'è ammesso
    Dim UtenteAmmesso As Boolean
    'Apertura del database
    Set DB = DBEngine.OpenDatabase(App.Path & "\Jukebox.mdb")
    'Viene avviata una scansione di tutti i file presenti all'interno del cd
    For I = 0 To ListaFile.ListCount
        'Se il file correntemente analizzato è quello contenente il numero ID utente,
        'allora
        If ListaFile.List(I) = "IDUtente.dat" Then
            'Apro il file contenente il numero utente
            Open ListaFile.Path & "IDUtente.dat" For Input As #1
                'Prelevo una riga dal file
                Line Input #1, Riga
                'Definisco la query di ricerca IDUtente all'interno dell'elenco utenti
                'permessi
                SQL = "SELECT * FROM Utenti WHERE ID = " & Riga
                'Viene eseguita la query di ricerca
                Set RS = DB.OpenRecordset(SQL)
                'Viene verificato se il codice utente è ammesso all'interno del jukebox
                While Not RS.EOF
                    'Salvo i rispettivi parametri
                    With InfoUtente
                        'Salvo il nome
                        .NomeUtente = RS("Utente")
                        'Salvo la password
                        .Password = RS("Password")
                        'Se al numero ID trovato corrisponde quello dell'amministratore, allora...
                        If RS("Utente") = "AMMINISTRATORE" Then
                            'Setto l'acceso come amministratore
                            .Accesso = Amministratore
                        'Altrimenti
                        Else
                            'Setto l'accesso come comune utente
                            .Accesso = utente
                        End If
                        'Setto il nome utente all'interno del rispettivo label
                        NomeUtente = InfoUtente.NomeUtente
                    End With
                    'Si passa ad analizzare il record successivo
                    RS.MoveNext
                Wend
                'Se il codice utente presente all'interno del cd non è ammesso, allora...
                If InfoUtente.Accesso = 0 Then
                    'Numero ID non ammesso
                    UtenteAmmesso = False
                'Altrimenti...
                Else
                    'Setto la variabile UtenteAmmesso a valore boleano True, in modo tale
                    'da far capire al programma che il file contenente il numero ID dell'utente
                    'è stato trovato
                    UtenteAmmesso = True
                End If
                'File contenente il numero utente trovato
                UtenteTrovato = True
            'Chiudo il file contenente il numero utente
            Close #1
        'Altrimenti se il file attualmente analizzato è quello contenente l'elenco
        'dei CD...
        ElseIf ListaFile.List(I) = "CDList.dat" Then
            'Apro il file contenente la lista dei CD preferiti
            Open ListaFile.Path & "CDList.dat" For Input As #1
                'Finchè il file non è finito
                While Not EOF(1)
                    'Prelevo l'ID di un CD dal file
                    Line Input #1, Riga
                    'Salvo l'ID del cd all'interno dell'array contenente i cd preferiti
                    CDPreferiti(IPref) = Val(Riga)
                    'Incremento l'indice dell'array
                    IPref = IPref + 1
                'Continua a ciclare
                Wend
            'Chiudo il file contenente la lista dei cd
            Close #1
        End If
    'Si passa ad analizzare il file successivo
    Next
    If UtenteTrovato = True Then
        'Imposto il colore verde
        UtenteValido.ForeColor = vbGreen
        'Setto la caption del label UtenteValido
        UtenteValido = "Utente valido!!!"
        'Setto il label descrizione
        Descrizione = "Prego, inserisca la sua password all'interno del campo sottostante."
    Else
        'Imposto il colore rosso
        UtenteValido.ForeColor = vbRed
        'Setto la caption del label UtenteValido
        UtenteValido = "Utente sconosciuto!!!"
    End If
End Sub

Private Sub Apertura_Database()
    'Dichiarazione di un indice temporaneo
    Dim I As Integer
    'Dichiaro un secondo indice temporaneo
    Dim K As Integer
    'Definizione della query di conteccio numero CD presenti all'interno
    'del jukebox
    SQL = "SELECT Count(*) AS NCD "
    SQL = SQL & "FROM CD"
    'Viene eseguita la query sopra definita
    Set RS = DB.OpenRecordset(SQL)
    'Viene salvato all'interno della varibile pubblica NCD il numero doi CD restituito
    'dalla query appena effettuata
    NCD = RS!NCD
    ICanzone = 1
    ReDim CD(1 To NCD + 1) As Mesh8
    ReDim Paramcd(1 To NCD + 1) As tparamcd
    ReDim Pos(1 To NCD + 1) As TPos
    'Ridefinisco la query in modo da prelevare dal database tutte le immmagini raffiguranti
    'i vari CD
    SQL = "SELECT Artista,ID,Texture,Titolo "
    SQL = SQL & "FROM CD ORDER BY Artista"
    'Eseguo la nuova Query
    Set RS = DB.OpenRecordset(SQL)
    'Effettuo un ciclo al fine di recuperare il percorso e il nome di tutte le immagini
    'raffiguranti i vari cd e passarli al filmato flash il quale utilizzerà tali
    'dati per creare il menù principale
    For I = 1 To NCD
        'Carico all'interno dell'array ArrayCD nella cella numero I, il percorso della
        'texture del CD corrente
        Paramcd(I).Texture = RS!Texture
        Paramcd(I).Album = RS!Titolo
        Paramcd(I).Artista = RS!Artista
        Paramcd(I).ID = RS!ID
        Paramcd(I).Lasciato = True
        'Sposto il cursore al record successivo
        RS.MoveNext
    Next
    Paramcd(NCD + 1).Lasciato = True
    'Definisco la query che mi permetterà di ricavare il numero di canzoni presenti all'interno
    'del JukeBox
    SQL = "SELECT SUM(NumeroTraccie) AS NCanzoni "
    SQL = SQL & "FROM CD"
    'Viene eseguita la query sopra definita
    Set RS = DB.OpenRecordset(SQL)
    'Salvo all'interno della Variabile NCanzoni il dato appena ricavato
    NCanzoni = RS!NCanzoni
    'Definisco la query che mi permetterà di ricavare il numero di artisti presenti all'interno
    'del JukeBox
    SQL = "SELECT DISTINCT Artista AS NArtisti "
    SQL = SQL & "FROM CD"
    'Viene eseguita la query sopra definita
    Set RS = DB.OpenRecordset(SQL)
    'Ciclo per il prelievo del numero di artisti presenti
    While Not RS.EOF
        'Incremento la variabile NArtisti
           NArtisti = NArtisti + 1
        'Si passa al record successivo
        RS.MoveNext
    Wend
    SQL = "SELECT * "
    SQL = SQL & "FROM CD ORDER BY Artista"
    'Viene eseguita la query sopra definita
    Set RS = DB.OpenRecordset(SQL)
    '------------------------------------------------------------------------------
    'Carico i titoli delle canzoni di tutti i cd all'interno dei rispettivi record
    '------------------------------------------------------------------------------
    For I = 1 To NCD
        For K = 1 To 30
            Paramcd(I).Canzone(K) = RS.Fields("Canzone" + CStr(K))
            Paramcd(I).IDCanzone(K) = ICanzone
            If RTrim(Paramcd(I).Canzone(K)) <> "\" Then
                ICanzone = ICanzone + 1
            End If
        Next
        'Passo al record successivo
        RS.MoveNext
    Next
End Sub

Private Sub Form_Terminate()
    'Rimostro il cursore del mouse
    ShowCursor True
End Sub

Private Sub Lettera_GotFocus(Index As Integer)
    'Imposto il colore di scrittura a giallo
    Lettera(Index).ForeColor = vbYellow
    'Nascondo il cursore della textbox selezionata
    HideCaret Lettera(Index).hwnd
End Sub

Private Sub Lettera_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Dichiaro un indice temporaneo
    Dim I As Integer
    'DIchiaro una variabile temporanea
    Dim tmp As String
    'Se il tasto premuto è lo 0, allora...
    If KeyCode = 96 Then
        'Viene verificato l'indice
        Select Case Index
        'Nel caso ci si sia spostati nella casella contenente lo spazio,allora...
        Case 26:
            'Aggiungo uno spazio
            NuovoNome = NuovoNome & " "
        'Nel caso ci si sia spostati nella casella addetta alla cancellazione,allora...
        Case 27:
            'Salvo all'interno della variabile temporanea il nome dell'utente privato
            'della sua ultima lettera
            tmp = Mid(NuovoNome, 1, Len(NuovoNome) - 1)
            'Riscrivo il nuovo nome utente modificato
            NuovoNome = tmp
        'Nel caso ci si sia spostati nella casella addetta al maiuscolo,allora...
        Case 28:
            If Lettera(28) = "minusc" Then
                'Converto in minuscolo tutte le lettere
                For I = 0 To 25
                    Lettera(I) = LCase(Lettera(I))
                Next
                'Imposto il nuovo testo
                Lettera(28) = "Maiusc"
            Else
                'Converto in maiuscolo tutte le lettere
                For I = 0 To 25
                    Lettera(I) = UCase(Lettera(I))
                Next
                'Imposto il nuovo testo
                Lettera(28) = "minusc"
            End If
        'Nel caso ci si sia spostati nella casella addetta al minuscolo,allora...
        Case 29:
            'Se il nome utente è stato impostato,allora...
            If NuovoNome <> "" Then
                'Salvo il nuovo nome definito
                ValNuovoNome = NuovoNome
                'Mostro il frame di definizione password nuovo utente
                FramePNU.Visible = True
            End If
        'In tutti gli altri casi
        Case Else:
            'Aggiungo la lettera selezionata al nuovo nome utente
            NuovoNome = NuovoNome & Lettera(Index)
        End Select
    ElseIf KeyCode = 97 Then
        On Error Resume Next
        Lettera(Index - 1).SetFocus
        If Err.Number <> 0 Then
            Lettera(29).SetFocus
        End If
    ElseIf KeyCode = 99 Then
        On Error Resume Next
        Lettera(Index + 1).SetFocus
        If Err.Number <> 0 Then
            Lettera(0).SetFocus
        End If
    ElseIf KeyCode = 27 Then
        ShowCursor True
        End
    End If
End Sub

Private Sub Lettera_LostFocus(Index As Integer)
    'Imposto il colore di scrittura a bianco
    Lettera(Index).ForeColor = vbWhite
End Sub

Private Sub Password_GotFocus()
    'Nascondo il cursore
    HideCaret Password.hwnd
End Sub

Private Sub Password_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 27:
        'Chiudo il programma
        End
    Case 96:
        'Viene richiamata la funzione addetta alla verifica della password inserita
        Verifica_Password
    End Select
    'Aggiorno il contenuto del campo Password
    'Password = ValPassword
End Sub

Private Sub Verifica_Password()
    'Se la password inserita è corretta, allora...
    If Password = InfoUtente.Password Then
        'Messaggio di OK
        VoceLogIN.Speak "Password OK!!!             Enjoy your music!"
        'Viene richiamata la funzione di Caricamento CD
        Apertura_Database
        'Mostro il form principale
        Main.Show
        'Chiudo questo form
        Unload Me
    'Altrimenti la password non è corretta, quindi...
    Else
        'Azzero il contenuto del campo password
        Password = ""
    End If
End Sub

Private Sub Correggi_Campo(Campo As TextBox)
    'Dichiaro una variabile temporanea
    Dim tmp As String
    On Error Resume Next
    tmp = Mid(Campo, 1, Len(Campo) - 1)
    BUG = False
    If Err.Number <> o Then
        tmp = ""
    End If
    Campo = tmp
End Sub

Private Sub verifica_Nuova_Password()
    'Se la nuova Password coincide con la nuova password di conferma,allora..
    If CNP(1) = CNP(2) Then
        'Salvo la Password
        ValPassword = CNP(1)
        'Richiamo la funzione addetta alla visualizzazione del messaggio prossima operazione
        Messaggio "Creazione Jukebox AccessCD", "Per completare l'operazione di creazione nuovo utente è necassario inserire, nell'apposito vano CD, un jukebox AccessCD vergine, in modo da poterle consentire tutti i suoi futuri accessi mediante lo stesso."
        'Lettura a voce del messaggio
        VoceLogIN.Speak "Please insert an empty jukebox access cd"
        'Viene aperto il vano CD
        SelectedDrive.EjectCD
        'Attivo il timer di controllo inserimento Jukebox AccessCD vergine
        TimerNewACD.Enabled = True
    'Altrimenti, si è verificato un errore, quindi...
    Else
        'Richiamo la funzione addetta alla visualizzazione di messaggi di errore
        MessaggioErrore "ATTENZIONE !!!", "Si è verificato un errore nella definizione della Password del nuovo utente!!! Si prega di verificare che la password di conferma coincida con la password, quindi spostarsi con i tasti direzionali dul pulsante di conferma e premere il tasto <OK>"
    End If
    
End Sub

Private Sub MessaggioErrore(Titolo As String, Testo As String)
    'Mostro il messaggio di errore
    ErrorMessage.Visible = True
    'Imposto il suo titolo
    ERTitolo = Titolo
    'Imposto il suo testo
    ERTesto = Testo
    'Imposto il focus sul pulsante di chiusura
    ERChiudi.SetFocus
End Sub

Private Sub Messaggio(Titolo As String, Testo As String)
    'Mostro il messaggio
    MessageFrame.Visible = True
    'Imposto il suo titolo
    TitoloMessaggio = Titolo
    'Imposto il suo testo
    TestoMessaggio = Testo
End Sub

Private Sub TimerNewACD_Timer()
    'Viene verificato se è stato inserito un CD e,in caso affermativo, viene controllato
    'se questo è vergine...
    On Error Resume Next
    Open SelectedDrive.DriveLetter + ":\testcd" For Output As #1
    Close #1
    'Se non si è verificato nessun errore, vuol dire che il cd è stato inserito,
    'quindi...
    If Err.Number <> 52 Then
        'Se si è verificato un errore con il CD,allora..
        If Err.Number = 75 Then
            'Viene nascosto il precedente messaggio
            MessageFrame.Visible = False
            'Viene visualizzato il messaggio di avvertitmento
            MessaggioErrore "ATTENZIONE !!!", "Il CD inserito non è vergine, oppure non è stato inserito alcun CD." + Chr(13) + "Si prega di inserire un jukebox AccessCD vergine, al fine di completare l'operazione di creazione nuovo utente."
            'Viene riaperto il vano CD
            SelectedDrive.EjectCD
        End If
    'Altrimenti, il cd è vuoto, quindi...
    Else
        'Viene nascosto il precedente messaggio
        MessageFrame.Visible = False
        'Chiamata alla funzione di creazione CD
        Crea_Jukebox_Access_CD
        'Disabilito il timer
        TimerNewACD.Enabled = False
    End If
End Sub

Private Sub Crea_Jukebox_Access_CD()
    'Dichiaro una variabile temporanea che avrà lo scopo di memorizzare l'ID del
    'nuovo utente
    Dim tmpID As Integer
    '--------------------------------------------
    ' Aggiunta del nuovo utente nel Database
    '--------------------------------------------
    'Apertura del database
    Set DB = DBEngine.OpenDatabase(App.Path & "\Jukebox.mdb")
    'Definizione della query di inserimento nuovo utente
    SQL = "INSERT INTO Utenti (Utente,Password,Data,Ora) VALUES ('" & ValNuovoNome & "','" & ValPassword & "','" & Date & "','" & Time & "')"
    'Viene eseguita la query sopra definita
    DB.Execute SQL
    'Definzisco una query al fine di ricercare il numero ID assegnato al nuovo utente
    SQL = "SELECT ID FROM Utenti WHERE Utente = '" & ValNuovoNome & "' AND Password = '" & ValPassword & "' AND Data = '" & Date & "'"
    'Eseguo la query di ricerca
    Set RS = DB.OpenRecordset(SQL)
    'Salvo il numero ID ritornato dalla query
    tmpID = RS("ID")
    'Creo il file contenente il numero ID del nuovo utente
    Open App.Path & "\IDUtente.dat" For Output As #1
        'Scrivo il numero ID sul file
        Write #1, tmpID
    'Chiudo il file temporaneo
    Close #1
    '--------------------------------------------
    ' Scrittura Jukebox AccessCD
    '--------------------------------------------
    'Assegno il file alla variabile Nero addetta alla sua identificazione
    FileID.Name = "IDUtente.dat"
    FileID.SourceFilePath = App.Path & "\IDUtente.dat"
    'Aggiungo il file alla cartella Root di nero
    Cartella.Files.Add FileID
    'Setto la modalità di scrittura del CD
    Traccia.BurnOptions = NERO_BURN_OPTION_USE_JOLIET Or NERO_BURN_OPTION_CREATE_ISO_FS
    'Setto la radice del cd
    Traccia.RootFolder = Cartella
    'Setto il nome del cd
    Traccia.Name = "Jukebox AccessCD"
    'Scrivo il CD
    SelectedDrive.BurnIsoAudioCD "", "", False, Traccia, Nothing, Nothing, NERO_BURN_FLAG_WRITE, 4, NERO_MEDIA_CD
    'Nacondo un eventuale residuo di messaggio di errore
    ErrorMessage.Visible = False
    'Viene visualizzato un messaggio di masterizzazione in corso
    Messaggio "Creazione Jukebox AccessCD", "Attendere prego!" + Chr(13) + "E' in corso la creazione del jukebox AccessCD per l'utente " + ValNuovoNome + "." + Chr(13) + "Il CD verrà automaticamente espulso al termine della procedura!"
End Sub
