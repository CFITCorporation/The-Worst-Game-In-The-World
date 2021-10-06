VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UltraRunner"
   ClientHeight    =   6225
   ClientLeft      =   165
   ClientTop       =   3930
   ClientWidth     =   9630
   Icon            =   "UltraRunner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9630
   Begin VB.HScrollBar Slider1 
      Height          =   255
      Left            =   6120
      Max             =   22
      Min             =   1
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   420
      Value           =   1
      Width           =   3015
   End
   Begin VB.OptionButton O1 
      BackColor       =   &H80000004&
      Caption         =   "II"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6960
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton S1 
      BackColor       =   &H80000004&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3000
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Z 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4440
      Top             =   0
   End
   Begin VB.Timer F 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   3480
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      DrawWidth       =   2
      FillColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   4740
      Left            =   480
      ScaleHeight     =   4710
      ScaleWidth      =   8625
      TabIndex        =   0
      Top             =   720
      Width           =   8655
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFDC&
         ForeColor       =   &H00800000&
         Height          =   1455
         Left            =   2280
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   3375
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Next Level "
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1020
            TabIndex        =   12
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lev1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "I-0"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1920
            TabIndex        =   11
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lev 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Level"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   960
            TabIndex        =   10
            Top             =   720
            Width           =   660
         End
         Begin VB.Label pas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "you've already passed"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   480
            TabIndex        =   9
            Top             =   480
            Width           =   2385
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Congratulations!"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1845
         End
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFC0&
         FillStyle       =   0  'Solid
         Height          =   4695
         Left            =   8520
         Top             =   0
         Width           =   120
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFC0&
         FillStyle       =   0  'Solid
         Height          =   4695
         Left            =   0
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Rhn 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   240
      End
      Begin VB.Shape Rhnn 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   270
         Left            =   120
         Shape           =   1  'Square
         Top             =   2040
         Width           =   270
      End
      Begin VB.Shape Fd 
         BorderColor     =   &H00008080&
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   8
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Fd 
         BorderColor     =   &H00008080&
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   7
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Fd 
         BorderColor     =   &H00008080&
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Fd 
         BorderColor     =   &H00008080&
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape Fd 
         BorderColor     =   &H00008080&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Fd 
         BorderColor     =   &H00008080&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Fd 
         BorderColor     =   &H00008080&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Fd 
         BorderColor     =   &H00008080&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Fd 
         BorderColor     =   &H00008080&
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   4080
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   44
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   43
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   42
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   41
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   40
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   39
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   38
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   37
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   36
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   35
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   34
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   33
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   32
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   31
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   30
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   29
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   28
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   27
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   26
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   25
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   24
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   23
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   22
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   21
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   20
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   19
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   0
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   18
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   17
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   16
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   15
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   14
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   13
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   12
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   11
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   10
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   9
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   8
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   7
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   6
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   5
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   4
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   3
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   2
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape zz 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   1
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape ck1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   8640
         Top             =   -120
         Width           =   15
      End
      Begin VB.Shape ck2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   4815
         Left            =   7680
         Top             =   0
         Width           =   975
      End
      Begin VB.Shape rk2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   4695
         Left            =   0
         Top             =   0
         Width           =   1560
      End
      Begin VB.Shape rk1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   15
         Left            =   0
         Top             =   0
         Width           =   0
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      Height          =   615
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   5580
      Width           =   1935
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   3840
      TabIndex        =   14
      Top             =   5640
      Width           =   885
   End
   Begin VB.Label Leve 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I-1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   4800
      TabIndex        =   13
      Top             =   5640
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5400
      TabIndex        =   4
      Top             =   0
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Style:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5400
      TabIndex        =   3
      Top             =   360
      Width           =   675
   End
   Begin VB.Label DC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Death Counter:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Dim r, L, U, D, j As Long
Dim i, Rr, Sp, Styles, Tot As Integer
Dim X0!, Y0!
Dim pi As Double
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Sub F_Timer()
j = j + 1
For i = 1 To 20
If i Mod 2 = 0 Then ' And zz(i).Top >= 180 Then
zz(i).Top = zz(i).Top - 60 - Sin(j) * 15
Else
zz(i).Top = zz(i).Top + 60 + Sin(j) * 15
End If
If i Mod 2 = 0 And zz(i).Top <= 0 Then
Z.Enabled = True
F.Enabled = False
End If
Next i
For i = 21 To 44
If i Mod 2 = 0 Then
If S1.Value = True Then
zz(i).Left = (Picture1.Width - ck2.Width + rk2.Width - zz(i).Width) / 2 + Sin((i - j / Sp - 21) * pi / 12) * Rr
zz(i).Top = (Picture1.Height - zz(i).Height) / 2 + Cos((i - j / Sp - 21) * pi / Styles) * Rr
Else
zz(i).Left = (Picture1.Width - ck2.Width + rk2.Width - zz(i).Width) / 2 + Cos((i - j / Sp - 21) * pi / 12) * Rr
zz(i).Top = (Picture1.Height - zz(i).Height) / 2 + Sin((i - j / Sp - 21) * pi / Styles) * Rr
End If
Else
zz(i).Left = (Picture1.Width - ck2.Width + rk2.Width - zz(i).Width) / 2 + Sin((i - j / Sp - 21) * pi / 12) * Rr * 2 / 3
zz(i).Top = (Picture1.Height - zz(i).Height) / 2 + Cos((i - j / Sp - 21) * pi / Styles) * Rr * 2 / 3
End If
Next i
For i = 1 To 8
If Fd(i).Tag = 0 Then
Fd(i).Left = (Picture1.Width - ck2.Width + rk2.Width - Fd(i).Width) / 2 + Sin((i - j / Sp / 16 - 21) * pi / 4) * Rr / 1.1
Fd(i).Top = (Picture1.Height - Fd(i).Height) / 2 + Cos((i - j / Sp / 16 - 21) * pi / 4) * Rr / 1.1
End If '
Next i
End Sub

Private Sub Form_Load()
Tot = 0
Styles = 24
Slider1.Value = 1
Sp = 8 '¡ðÐý×ªËÙ¶È
Rr = 2160 '8Çò×é³É¡ð°ë¾¶
pi = 3.14159265358979
r = 39
L = 37
U = 38
D = 40
For i = 1 To 8
Fd(i).Tag = 0
Fd(i).BorderStyle = 0
Fd(i).FillColor = &HFF00FF
Next i
For i = 1 To 20
If i Mod 2 = 0 Then
zz(i).Left = rk2.Width + zz(i).Width + Int((Picture1.Width - rk2.Width - ck2.Width) / 20) * (i - 1.85)
zz(i).Top = 0
Else
zz(i).Left = rk2.Width + zz(i).Width + Int((Picture1.Width - rk2.Width - ck2.Width) / 20) * (i - 1.8)
zz(i).Top = Picture1.Height - zz(i).Height
End If
Next i
For i = 21 To 44
If i Mod 2 = 0 Then
zz(i).Left = (Picture1.Width - ck2.Width + rk2.Width - zz(i).Width) / 2 + Sin((i - 21) * pi / 12) * Rr
zz(i).Top = (Picture1.Height - zz(i).Height) / 2 + Cos((i - 21) * pi / 12) * Rr
Else
zz(i).Left = (Picture1.Width - ck2.Width + rk2.Width - zz(i).Width) / 2 + Sin((i - 21) * pi / 12) * Rr * 3 / 2
zz(i).Top = (Picture1.Height - zz(i).Height) / 2 + Cos((i - 21) * pi / 12) * Rr * 2 / 3
End If
Next i
Z.Enabled = True
End Sub

Private Sub Label5_Click()
Frame1.Visible = False
Rhn.Left = 120
Rhn.Top = 2160
If Slider1.Value = Slider1.Max And S1.Value = True Then
Slider1.Value = 1
O1.Value = True
ElseIf Slider1.Value = Slider1.Max And O1.Value = True Then
Slider1.Value = 1
S1.Value = True
DC.Caption = 0
Else
Slider1.Value = Slider1.Value + 1
End If
Styles = 25 - Slider1.Value
If S1.Value = True Then
Leve.Caption = Replace(S1.Caption + "-" + Str(Slider1.Value), " ", "")
Else
Leve.Caption = Replace(O1.Caption + "-" + Str(Slider1.Value), " ", "")
End If
End Sub

Private Sub O1_Click()
Picture1.Cls
If S1.Value = True Then
Leve.Caption = Replace(S1.Caption + "-" + Str(Slider1.Value), " ", "")
Else
Leve.Caption = Replace(O1.Caption + "-" + Str(Slider1.Value), " ", "")
End If
End Sub

'38
'40
'37
'39
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print (KeyCode)
If KeyCode = L And Rhn.Left > 0 Then '((Rhn.Top <= rk1.Height + Rhn.Height And Rhn.Left >= rk1.Width - 120) Or (Rhn.Top > rk1.Top - Rhn.Height And Rhn.Left > rk1.Width) Or (Rhn.Top >= rk1.Height - 120 And Rhn.Top <= rk1.Top - Rhn.Height + 120 And Rhn.Left > 0)) Then
Rhn.Left = Rhn.Left - 90
ElseIf KeyCode = r And Rhn.Left < -Rhn.Width + Picture1.Width Then '((Rhn.Top <= ck1.Height + Rhn.Height And Rhn.Left < ck1.Left - 120 - Rhn.Width) Or (Rhn.Top > ck1.Top - Rhn.Height And Rhn.Left < ck1.Left - 120 - Rhn.Width) Or (Rhn.Top >= ck1.Height - 120 And Rhn.Top <= ck1.Top - Rhn.Height + 120 And Rhn.Left < Picture1.Width - Rhn.Width)) Then
Rhn.Left = Rhn.Left + 90
ElseIf KeyCode = U And (Rhn.Top > 0) Then 'Rhn.Left >= rk1.Width - 120 And (Rhn.Left <= ck1.Left - Rhn.Width) And
Rhn.Top = Rhn.Top - 90
ElseIf KeyCode = D And (Rhn.Top < -Rhn.Height + Picture1.Height) Then 'Rhn.Left >= rk1.Width - 90 And (Rhn.Left <= ck1.Left - Rhn.Width) And
Rhn.Top = Rhn.Top + 90
End If
Z.Interval = 30
F.Interval = 30
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Abs(X - Rhn.Left) < Rhn.Width And Abs(Y - Rhn.Top) < Rhn.Height Then
'X0 = X
'Y0 = Y
'End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 And Abs(X - Rhn.Left) < Rhn.Width And Abs(Y - Rhn.Top) < Rhn.Height Then
'Rhn.Left = Rhn.Left + X - X0
'Rhn.Top = Rhn.Top + Y - Y0
'Debug.Print Rhn.Left, X, X0
'End If
End Sub

Private Sub Rhn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Abs(X - Rhn.Left) < Rhn.Width And Abs(Y - Rhn.Top) < Rhn.Height Then
X0 = X
Y0 = Y
Rhnn.Left = Rhn.Left
Rhnn.Top = Rhn.Top
'End If
End Sub

Private Sub Rhn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ' And Abs(X - Rhn.Left) < Rhn.Width And Abs(Y - Rhn.Top) < Rhn.Height Then
'If X - X0 <= 0 And Rhn.Left > 0 Then '((Rhn.Top <= rk1.Height + Rhn.Height And Rhn.Left >= rk1.Width - 120) Or (Rhn.Top > rk1.Top - Rhn.Height And Rhn.Left > rk1.Width) Or (Rhn.Top >= rk1.Height - 120 And Rhn.Top <= rk1.Top - Rhn.Height + 120 And Rhn.Left > 0)) Then
'And ((Rhn.Top <= rk1.Height + Rhn.Height And Rhn.Left > rk1.Width + 120) Or (Rhn.Top > rk1.Top - Rhn.Height And Rhn.Left > rk1.Width) Or (Rhn.Top >= rk1.Height - 120 And Rhn.Top <= rk1.Top - Rhn.Height + 120 And Rhn.Left > 0)) Then
'Rhn.Left = Rhn.Left + X - X0
'Rhn.Top = Rhn.Top + Y - Y0
'Debug.Print "L", Rhn.Left, X, X0
'ElseIf X - X0 >= 0 And Rhn.Left < -Rhn.Width + Picture1.Width Then 'And ((Rhn.Top <= ck1.Height + Rhn.Height And Rhn.Left < ck1.Left - 120 - Rhn.Width) Or (Rhn.Top > ck1.Top - Rhn.Height And Rhn.Left < ck1.Left - 120 - Rhn.Width) Or (Rhn.Top >= ck1.Height - 120 And Rhn.Top <= ck1.Top - Rhn.Height + 120 And Rhn.Left < Picture1.Width - Rhn.Width)) Then
'Rhn.Left = Rhn.Left + X - X0
'Rhn.Top = Rhn.Top + Y - Y0
'Debug.Print "R", Rhn.Left, X, X0
If Rhn.Left < -Rhn.Width + Picture1.Width And Rhn.Left < -Rhn.Width + Picture1.Width + 120 And (Rhn.Top > -120) And (Rhn.Top < -Rhn.Height + Picture1.Height + 120) And Rhn.Left >= 0 Then 'And (Rhn.Left >= rk1.Width - 120 And (Rhn.Left <= ck1.Left - Rhn.Width) And Rhn.Top > -120) Then
'Rhn.Left = Rhn.Left + X - X0If Y - Y0 <= 0 And
'Rhn.Top = Rhn.Top + Y - Y0
'Debug.Print "U", Rhn.Top, Y, Y0
'ElseIf Y - Y0 >= 0 And (Rhn.Top < -Rhn.Height + Picture1.Height) Then  '(Rhn.Left >= rk1.Width - 90 And (Rhn.Left <= ck1.Left - Rhn.Width) And Rhn.Top < -Rhn.Height + Picture1.Height) Then
Rhn.Left = Rhn.Left + X - X0
Rhn.Top = Rhn.Top + Y - Y0
'Debug.Print "D", Rhn.Top, Y, Y0
End If

Z.Interval = 20 '8
F.Interval = 20 '8
ElseIf Button = 0 Then
Z.Interval = 40
F.Interval = 40
End If
End Sub

Private Sub Rhn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rhnn.Left = Rhn.Left
Rhnn.Top = Rhn.Top
If Rhnn.Top > Picture1.Height - Rhn.Height Then
Rhn.Top = -Rhn.Height + Picture1.Height - 60
ElseIf Rhnn.Left > Picture1.Width - Rhn.Width Then
Rhn.Left = -Rhn.Width + Picture1.Width - 120
ElseIf Rhnn.Top < 0 Then
Rhn.Top = 0
ElseIf Rhnn.Left < 0 Then
Rhn.Left = 120
End If
Debug.Print 1
End Sub

Private Sub S1_Click()
Picture1.Cls
If S1.Value = True Then
Leve.Caption = Replace(S1.Caption + "-" + Str(Slider1.Value), " ", "")
Else
Leve.Caption = Replace(O1.Caption + "-" + Str(Slider1.Value), " ", "")
End If
End Sub

Private Sub Slider1_Change()
Picture1.Cls
Rhn.Left = 120
Rhn.Top = 2160
Styles = 25 - Slider1.Value
If S1.Value = True Then
Leve.Caption = Replace(S1.Caption + "-" + Str(Slider1.Value), " ", "")
Else
Leve.Caption = Replace(O1.Caption + "-" + Str(Slider1.Value), " ", "")
End If
Tot = 0
Dim j As Integer
For j = 1 To 8
Fd(j).Tag = 0
Fd(j).Visible = True
Next j
End Sub

Private Sub Slider1_Click()
Picture1.Cls
Rhn.Left = 120
Rhn.Top = 2160
Styles = 25 - Slider1.Value
If S1.Value = True Then
Leve.Caption = Replace(S1.Caption + "-" + Str(Slider1.Value), " ", "")
Else
Leve.Caption = Replace(O1.Caption + "-" + Str(Slider1.Value), " ", "")
End If
Tot = 0
Dim j As Integer
For j = 1 To 8
Fd(j).Tag = 0
Fd(j).Visible = True
Next j
End Sub

Private Sub Slider1_Scroll()
Picture1.Cls
Rhn.Left = 120
Rhn.Top = 2160
Styles = 25 - Slider1.Value
If S1.Value = True Then
Leve.Caption = Replace(S1.Caption + "-" + Str(Slider1.Value), " ", "")
Else
Leve.Caption = Replace(O1.Caption + "-" + Str(Slider1.Value), " ", "")
End If
Tot = 0
Dim j As Integer
For j = 1 To 8
Fd(j).Tag = 0
Fd(j).Visible = True
Next j
End Sub

Private Sub Timer1_Timer()
If Rhnn.Top > Picture1.Height - Rhn.Height Then
Rhn.Top = -Rhn.Height + Picture1.Height - 60
ElseIf Rhnn.Left > Picture1.Width - Rhn.Width - 120 Then
Rhn.Left = -Rhn.Width + Picture1.Width - 120
ElseIf Rhnn.Top < 0 Then
Rhn.Top = 0
ElseIf Rhnn.Left < 0 Then
Rhn.Left = 120
End If
'0,2040Dim j As Integer
Dim j As Integer
For i = 1 To 44
Picture1.PSet (zz(i).Left + Int(zz(i).Width / 2), zz(i).Top + Int(zz(i).Height / 2)), &HFFFF00
If Abs(zz(i).Left - Rhn.Left) < 150 And Abs(zz(i).Top - Rhn.Top) < 150 Then
Tot = 0

For j = 1 To 8
Fd(j).Tag = 0
Fd(j).Visible = True
Next j
Rhn.Left = 120
Rhn.Top = 2160
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
DC.Caption = DC.Caption + 1
Exit For
ElseIf Rhn.Left >= ck2.Left And Tot = 8 Then
Tot = 0

For j = 1 To 8
Fd(j).Tag = 0
Fd(j).Visible = True
Next j

Frame1.Visible = True
pas.Left = Frame1.Width / 2 - pas.Width / 2
lev.Left = Frame1.Width / 2 - lev.Width - 120
lev1.Left = Frame1.Width / 2 + 120
If S1.Value = True Then
lev1.Caption = Replace(S1.Caption + "-" + Str(Slider1.Value), " ", "")
Else
lev1.Caption = Replace(O1.Caption + "-" + Str(Slider1.Value), " ", "")
End If
End If
Next i
For i = 1 To 8
If Abs(Fd(i).Left - Rhn.Left) < 150 And Abs(Fd(i).Top - Rhn.Top) < 150 Then
Fd(i).Visible = False
Fd(i).Tag = 1
Fd(i).Left = 0
Fd(i).Top = 0
Tot = Tot + 1
End If
Next i
End Sub

Private Sub Timer2_Timer()
Rhnn.Left = Rhn.Left
Rhnn.Top = Rhn.Top
End Sub

Private Sub Z_Timer()
j = j + 1
For i = 1 To 20
If i Mod 2 = 0 Then 'And zz(i).Top <= Picture1.Height - zz(i).Height - 180 Then
zz(i).Top = zz(i).Top + 60 + Cos(j) * 15
Else
zz(i).Top = zz(i).Top - 60 - Cos(j) * 15
End If

If i Mod 2 = 0 And zz(i).Top >= Picture1.Height - zz(i).Height Then
Z.Enabled = False
F.Enabled = True
End If
Next i
For i = 21 To 44
If i Mod 2 = 0 Then
If S1.Value = True Then
zz(i).Left = (Picture1.Width - ck2.Width + rk2.Width - zz(i).Width) / 2 + Sin((i - j / Sp - 21) * pi / 12) * Rr
zz(i).Top = (Picture1.Height - zz(i).Height) / 2 + Cos((i - j / Sp - 21) * pi / Styles) * Rr
Else
zz(i).Left = (Picture1.Width - ck2.Width + rk2.Width - zz(i).Width) / 2 + Cos((i - j / Sp - 21) * pi / 12) * Rr
zz(i).Top = (Picture1.Height - zz(i).Height) / 2 + Sin((i - j / Sp - 21) * pi / Styles) * Rr
End If
Else
zz(i).Left = (Picture1.Width - ck2.Width + rk2.Width - zz(i).Width) / 2 + Sin((i - j / Sp - 21) * pi / 12) * Rr * 2 / 3
zz(i).Top = (Picture1.Height - zz(i).Height) / 2 + Cos((i - j / Sp - 21) * pi / Styles) * Rr * 2 / 3
End If
Next i
For i = 1 To 8
If Fd(i).Tag = 0 Then
Fd(i).Left = (Picture1.Width - ck2.Width + rk2.Width - Fd(i).Width) / 2 + Sin((i - j / Sp / 16 - 21) * pi / 4) * Rr / 1.1
Fd(i).Top = (Picture1.Height - Fd(i).Height) / 2 + Cos((i - j / Sp / 16 - 21) * pi / 4) * Rr / 1.1
End If '
Next i
End Sub
