VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lem"
   ClientHeight    =   6795
   ClientLeft      =   1905
   ClientTop       =   660
   ClientWidth     =   7260
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   Begin VB.Timer tmrAdd 
      Left            =   3780
      Top             =   3600
   End
   Begin VB.Timer tmrPhysics 
      Left            =   3060
      Top             =   3180
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Walker"
      Height          =   375
      Index           =   0
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   5340
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Play God (Lemming Add)"
      Height          =   495
      Left            =   5460
      TabIndex        =   58
      Top             =   6240
      Width           =   1755
   End
   Begin VB.CommandButton cmdExplode 
      Caption         =   "Play God (Explode All)"
      Height          =   495
      Left            =   3660
      TabIndex        =   57
      Top             =   6240
      Width           =   1755
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Play God (Pencil)"
      Height          =   495
      Index           =   9
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   6240
      Width           =   1755
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Play God (Eraser)"
      Height          =   495
      Index           =   10
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   6240
      Value           =   -1  'True
      Width           =   1755
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Builder"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5820
      Width           =   855
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Miner"
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   5820
      Width           =   855
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Floater"
      Height          =   375
      Index           =   7
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5820
      Width           =   855
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Exploder"
      Height          =   375
      Index           =   6
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   5820
      Width           =   855
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Digger"
      Height          =   375
      Index           =   5
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   5820
      Width           =   855
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Climber"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5820
      Width           =   855
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Blocker"
      Height          =   375
      Index           =   2
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   5820
      Width           =   855
   End
   Begin VB.OptionButton optJob 
      Caption         =   "Basher"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5820
      Width           =   855
   End
   Begin VB.Frame fraFrames 
      Caption         =   "Frames"
      Height          =   3675
      Left            =   6900
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
      Begin VB.PictureBox picFrameFloat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   1980
         Picture         =   "frmMain.frx":0CCA
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picFrameFloat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   2340
         Picture         =   "frmMain.frx":148C
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   61
         Top             =   240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picFrameFloat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   2
         Left            =   2700
         Picture         =   "frmMain.frx":1C4E
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   60
         Top             =   240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picFrameFloat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   3
         Left            =   3000
         Picture         =   "frmMain.frx":2398
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   59
         Top             =   240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picPencil 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1680
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   56
         Top             =   2880
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picNumber 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   480
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   53
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picNumber 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   5
         Left            =   1020
         Picture         =   "frmMain.frx":2A6A
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   44
         Top             =   2460
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox picNumber 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   780
         Picture         =   "frmMain.frx":2BEC
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   43
         Top             =   2460
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox picNumber 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   3
         Left            =   540
         Picture         =   "frmMain.frx":2DAE
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   42
         Top             =   2460
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox picNumber 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   2
         Left            =   300
         Picture         =   "frmMain.frx":2F30
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   10
         TabIndex        =   41
         Top             =   2460
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.PictureBox picNumber 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   1
         Left            =   60
         Picture         =   "frmMain.frx":30B2
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   4
         TabIndex        =   40
         Top             =   2460
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.PictureBox picClear 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1140
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   39
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   7
         Left            =   420
         Picture         =   "frmMain.frx":3178
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   38
         Top             =   1500
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   6
         Left            =   60
         Picture         =   "frmMain.frx":392A
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   37
         Top             =   1500
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   5
         Left            =   1800
         Picture         =   "frmMain.frx":40DC
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   36
         Top             =   1020
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   4
         Left            =   1500
         Picture         =   "frmMain.frx":47AE
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   35
         Top             =   1020
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   3
         Left            =   1140
         Picture         =   "frmMain.frx":4E10
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   34
         Top             =   1020
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   2
         Left            =   780
         Picture         =   "frmMain.frx":55C2
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   33
         Top             =   1020
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   1
         Left            =   420
         Picture         =   "frmMain.frx":5D74
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   32
         Top             =   1020
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   0
         Left            =   60
         Picture         =   "frmMain.frx":6526
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   31
         Top             =   1020
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   8
         Left            =   780
         Picture         =   "frmMain.frx":6CD8
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   30
         Top             =   1500
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   9
         Left            =   1140
         Picture         =   "frmMain.frx":748A
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   29
         Top             =   1500
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   10
         Left            =   1500
         Picture         =   "frmMain.frx":7C3C
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   24
         TabIndex        =   28
         Top             =   1500
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   11
         Left            =   60
         Picture         =   "frmMain.frx":845E
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   24
         TabIndex        =   27
         Top             =   1980
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   12
         Left            =   480
         Picture         =   "frmMain.frx":8C80
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   24
         TabIndex        =   26
         Top             =   1980
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   13
         Left            =   900
         Picture         =   "frmMain.frx":94A2
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   26
         TabIndex        =   25
         Top             =   1980
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   14
         Left            =   1320
         Picture         =   "frmMain.frx":9DA4
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   28
         TabIndex        =   24
         Top             =   1980
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox picFrameDig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   15
         Left            =   1800
         Picture         =   "frmMain.frx":A716
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   23
         Top             =   1980
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picFrameLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   1440
         Picture         =   "frmMain.frx":AEC8
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   11
         TabIndex        =   22
         Top             =   540
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.PictureBox picFrameLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   1200
         Picture         =   "frmMain.frx":B1DA
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   21
         Top             =   540
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picFrameLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   1020
         Picture         =   "frmMain.frx":B53C
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   11
         TabIndex        =   20
         Top             =   540
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.PictureBox picFrameLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   840
         Picture         =   "frmMain.frx":B84E
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   19
         Top             =   540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox picFrameLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   660
         Picture         =   "frmMain.frx":BAC0
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   11
         TabIndex        =   18
         Top             =   540
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.PictureBox picFrameLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   420
         Picture         =   "frmMain.frx":BDD2
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   17
         Top             =   540
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picFrameLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   240
         Picture         =   "frmMain.frx":C134
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   11
         TabIndex        =   16
         Top             =   540
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.PictureBox picFrameLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   60
         Picture         =   "frmMain.frx":C446
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   15
         Top             =   540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox picFrameRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   60
         Picture         =   "frmMain.frx":C6B8
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   14
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picFrameRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   300
         Picture         =   "frmMain.frx":CA1A
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   13
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picFrameRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   540
         Picture         =   "frmMain.frx":CD7C
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   12
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picFrameRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   780
         Picture         =   "frmMain.frx":D0DE
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picFrameRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   1020
         Picture         =   "frmMain.frx":D440
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   10
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picFrameRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   1260
         Picture         =   "frmMain.frx":D7A2
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   9
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picFrameRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   1500
         Picture         =   "frmMain.frx":DB04
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picFrameRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   1740
         Picture         =   "frmMain.frx":DE66
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdAddMinus 
      Caption         =   "-"
      Height          =   315
      Left            =   5820
      TabIndex        =   5
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton cmdAddPlus 
      Caption         =   "+"
      Height          =   315
      Left            =   5820
      TabIndex        =   4
      Top             =   5040
      Width           =   255
   End
   Begin VB.PictureBox picBoard 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4815
      Left            =   600
      MouseIcon       =   "frmMain.frx":E1C8
      Picture         =   "frmMain.frx":EE92
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
      Width           =   6060
   End
   Begin VB.Label lblAddInterval 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1500"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   6120
      TabIndex        =   3
      Top             =   5160
      Width           =   1020
   End
   Begin VB.Label lblLem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lemming #:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   60
      TabIndex        =   2
      Top             =   5340
      Width           =   1305
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "count:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   60
      TabIndex        =   1
      Top             =   5040
      Width           =   690
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ******************************
' Lemmings trial by Peter Rosconi
' Lines of code 184 (excluding
' comments and extra carriage returns)
' ******************************
Option Explicit

' ******************************
' BitBlt function
' ******************************
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' ******************************
' Fake timer objects
' ******************************
'Private WithEvents tmrPhysics As ccrpTimer
'Private WithEvents tmrAdd As ccrpTimer

' ******************************
' The attributes of each lemming
' ******************************
Private Type Lemming
    Alive As Boolean
    ExplodeStart As Double
    FallFrame As Integer
    Frame As Integer
    Job As Integer
    Left As Integer
    PushX As Single
    PushY As Single
    Top As Integer
    Walking As Boolean
    Width As Integer
End Type

' ******************************
' Lemming environment constants
' ******************************
Private Const LEM_WIDTH As Integer = 11
Private Const LEM_HEIGHT As Integer = 20
Private Const LEM_GRAVITY As Single = 1
Private Const LEM_STEP As Integer = 2
Private Const LEM_SPLAT As Integer = 15

' ******************************
' Lemming job constants
' ******************************
Private Const LEM_JOB_NOTHING As Integer = 0
Private Const LEM_JOB_BASHER As Integer = 1
Private Const LEM_JOB_BLOCKER As Integer = 2
Private Const LEM_JOB_BUILDER As Integer = 3
Private Const LEM_JOB_CLIMBER As Integer = 4
Private Const LEM_JOB_DIGGER As Integer = 5
Private Const LEM_JOB_EXPLODER As Integer = 6
Private Const LEM_JOB_FLOATER As Integer = 7
Private Const LEM_JOB_MINER As Integer = 8

' ******************************
' Lemming environment variables
' ******************************
Private lemAll() As Lemming
Private intJob As Integer
Private intStartLeft As Integer
Private intStartTop As Integer
Private intLemmings As Integer
Private intLemming As Integer
Private sngMouseX As Single
Private sngMouseY As Single

Private Sub cmdAdd_Click()
    intLemmings = intLemmings + 1
    Call addLemming
End Sub

Private Sub cmdAddMinus_Click()
    tmrAdd.Interval = tmrAdd.Interval - 10
    lblAddInterval.Caption = tmrAdd.Interval
End Sub

Private Sub cmdAddPlus_Click()
    tmrAdd.Interval = tmrAdd.Interval + 10
    lblAddInterval.Caption = tmrAdd.Interval
End Sub

Private Sub cmdExplode_Click()
    Dim i As Integer
    
    intLemmings = UBound(lemAll())
    For i = 0 To UBound(lemAll())
        lemAll(i).Job = LEM_JOB_EXPLODER
    Next
End Sub

Private Sub Form_Load()
    Call mciSendString("open " & App.Path & "\midi\cancanp.mid type sequencer alias midi", 0&, 0, 0)
    Call mciSendString("play midi", 0&, 0, 0)

    ' ******************************
    ' Declare the ccrpTimers
    ' ******************************
    'Set tmrPhysics = New ccrpTimer
    'Set tmrAdd = New ccrpTimer
   
    With tmrPhysics
        '.EventType = TimerPeriodic
        .Interval = 75
        .Enabled = True
    End With
    
    With tmrAdd
        '.EventType = TimerPeriodic
        .Interval = 1500
        .Enabled = True
    End With
    
    ' ******************************
    ' Set the starting positions and
    ' the amount of lemmings to populate
    ' ******************************
    intStartLeft = 150
    intStartTop = 0
    intLemmings = 200
    
    ' ******************************
    ' Redimension the lemmings and add the first one
    ' ******************************
    ReDim lemAll(0) As Lemming
    Call addLemming(True)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call mciSendString("close midi", 0&, 0, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mciSendString("close midi", 0&, 0, 0)
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub optJob_Click(Index As Integer)
    intJob = Index
End Sub

Private Sub picBoard_Click()
    If Not intLemming = -1 Then
        lemAll(intLemming).Job = intJob
    End If
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    sngMouseX = X
    sngMouseY = Y
    
    If Button = vbLeftButton And intJob = 10 Then
        Call BitBlt(picBoard.HDC, X, Y, LEM_WIDTH, LEM_HEIGHT, picClear.HDC, 0, 0, vbSrcCopy)
        picBoard.Refresh
    End If
    
    If Button = vbLeftButton And intJob = 9 Then
        Call BitBlt(picBoard.HDC, X, Y, LEM_WIDTH, LEM_HEIGHT, picPencil.HDC, 0, 0, vbSrcCopy)
        picBoard.Refresh
    End If
    
    For i = 0 To UBound(lemAll())
        If lemAll(i).Alive = False Then GoTo next_lem
        If X > lemAll(i).Left And X < lemAll(i).Left + LEM_WIDTH And Y > lemAll(i).Top And Y < lemAll(i).Top + LEM_HEIGHT Then
            ' ******************************
            ' Over lemming i
            ' ******************************
            intLemming = i
           
            lblLem.Caption = "Lemming #: " & intLemming & " | Job: " & optJob(lemAll(i).Job).Caption
            picBoard.MousePointer = vbCustom
            Exit Sub
        End If
next_lem:
    Next
    
    intLemming = -1
    picBoard.MousePointer = vbDefault
End Sub

Private Sub tmrAdd_Timer()
    ' ******************************
    ' Incremently add to the lemming population
    ' ******************************
    Call addLemming
End Sub

Private Sub tmrPhysics_Timer()
    Dim i As Integer
    Dim n As Integer
    Dim intPush As Integer
    Dim intPrevTop As Integer
    Dim blnWalk As Boolean
    
    ' Debug.Print (1000 / Milliseconds) ' Frames per second
    
    ' ******************************
    ' Clear the previous lemming so
    ' they don't interfere with each other
    ' ******************************
    For i = 0 To UBound(lemAll())
       Call drawLemming(i)
    Next
    
    ' ******************************
    ' Go through each alive lemming
    ' ******************************
    For i = 0 To UBound(lemAll())
        If lemAll(i).Alive = False Then GoTo next_lem
        If lemAll(i).Job = LEM_JOB_BLOCKER Then GoTo next_lem
        If lemAll(i).Job = LEM_JOB_DIGGER Then
            If lemAll(i).Frame / 5 = Int(lemAll(i).Frame / 5) Then
                lemAll(i).Top = lemAll(i).Top + LEM_STEP / 2
                If checkFeet(i) = False Then lemAll(i).Job = LEM_JOB_NOTHING
            End If
            GoTo next_lem
        End If
        
        ' ******************************
        ' If falling then don't move left or right
        ' ******************************
        If lemAll(i).FallFrame <= LEM_STEP Then lemAll(i).Left = lemAll(i).Left + lemAll(i).PushX
        
        ' ******************************
        ' Always try to fall
        ' ******************************
        lemAll(i).Top = lemAll(i).Top + lemAll(i).PushY
        If lemAll(i).PushY > (LEM_SPLAT / 5) And lemAll(i).Job = LEM_JOB_FLOATER Then
            lemAll(i).PushY = (LEM_SPLAT / 5)
        End If
        
        ' ******************************
        ' Check to see if you are on ground
        ' ******************************
        blnWalk = checkFeet(i)
        
        ' ******************************
        ' Not walking, but falling
        ' ******************************
        lemAll(i).Walking = False
        lemAll(i).FallFrame = lemAll(i).FallFrame + 1

        If blnWalk = True Then
            If lemAll(i).PushY > LEM_SPLAT Then
                Call killLemming(i, False, "thunk")
                GoTo next_lem
            End If
            ' ******************************
            ' Not falling, but walking
            ' ******************************
            lemAll(i).Walking = True
            lemAll(i).FallFrame = 0
            lemAll(i).PushY = LEM_STEP
            
            ' ******************************
            ' Check if in ground and push up
            ' ******************************
            intPush = 0
            intPrevTop = lemAll(i).Top
            Do
startOver:
                lemAll(i).Top = lemAll(i).Top - 1
                intPush = intPush + 1
                
                ' ******************************
                ' Turn around
                ' ******************************
                If intPush > LEM_HEIGHT Then
                    lemAll(i).PushX = lemAll(i).PushX * -1
                    lemAll(i).Left = lemAll(i).Left + lemAll(i).PushX
                    lemAll(i).Top = intPrevTop
                    Exit Do
                End If
                
                ' ******************************
                ' Check feet
                ' ******************************
                If checkFeet(i) = True Then GoTo startOver
            Loop Until picBoard.Point(lemAll(i).Left, lemAll(i).Top) = RGB(0, 0, 0)

        End If
        
        ' ******************************
        ' Always add to gravity
        ' ******************************
        lemAll(i).PushY = lemAll(i).PushY + LEM_GRAVITY
next_lem:
    Next

    ' ******************************
    ' Draw the lemmings last so they
    ' don't interfere with each other
    ' ******************************
    For i = 0 To UBound(lemAll())
        Call drawLemming(i, True)
    Next
    Call picBoard_MouseMove(0, 0, sngMouseX, sngMouseY)
End Sub

' ******************************
' This sub draws the lemming
' ******************************
Private Sub drawLemming(ByVal Index As Integer, Optional ByVal blnClear As Boolean)
    Dim intTimeLeft As Integer
    Dim pic As PictureBox
    
    If lemAll(Index).Alive = False Then Exit Sub
    If blnClear Then
        lemAll(Index).Frame = lemAll(Index).Frame + 1
        
        ' ******************************
        ' Which lemming to draw
        ' ******************************
        If lemAll(Index).Job = LEM_JOB_DIGGER Then
            If lemAll(Index).Frame > picFrameDig.UBound Then lemAll(Index).Frame = 0
            Set pic = picFrameDig(lemAll(Index).Frame)
            
            ElseIf lemAll(Index).Job = LEM_JOB_FLOATER And lemAll(Index).PushY > LEM_SPLAT / 5 Then
            If lemAll(Index).Frame > picFrameFloat.UBound Then lemAll(Index).Frame = 0
            Set pic = picFrameFloat(lemAll(Index).Frame)
            
            ElseIf lemAll(Index).PushX > 0 Then
            If lemAll(Index).Frame > picFrameRight.UBound Then lemAll(Index).Frame = 0
            Set pic = picFrameRight(lemAll(Index).Frame)
            
            Else
            If lemAll(Index).Frame > picFrameLeft.UBound Then lemAll(Index).Frame = 0
            Set pic = picFrameLeft(lemAll(Index).Frame)
        
        End If
        
        Call BitBlt(picBoard.HDC, lemAll(Index).Left, lemAll(Index).Top, LEM_WIDTH, LEM_HEIGHT, pic.HDC, 0, 0, vbSrcPaint)
        lemAll(Index).Width = pic.Width
        
        If lemAll(Index).Job = LEM_JOB_EXPLODER Then
            If lemAll(Index).ExplodeStart = 0 Then lemAll(Index).ExplodeStart = Timer
            intTimeLeft = 5 - Int(Timer - lemAll(Index).ExplodeStart)
            If intTimeLeft > 5 Then intTimeLeft = 5
            
            If intTimeLeft <= 0 Then
                Call killLemming(Index)
                GoTo clearLemming
            End If
            
            Call BitBlt(picBoard.HDC, lemAll(Index).Left, lemAll(Index).Top, LEM_WIDTH, LEM_HEIGHT, picNumber(intTimeLeft).HDC, 0, 0, vbSrcPaint)
        End If
        
        Else
        
        If lemAll(Index).Job = LEM_JOB_DIGGER Then
            Call BitBlt(picBoard.HDC, lemAll(Index).Left - 4, lemAll(Index).Top - 5, LEM_WIDTH + 8, LEM_HEIGHT + 5, picClear.HDC, 0, 0, vbSrcCopy)
        ElseIf lemAll(Index).Job = LEM_JOB_FLOATER And lemAll(Index).PushY > LEM_SPLAT / 5 Then
            Call BitBlt(picBoard.HDC, lemAll(Index).Left - 4, lemAll(Index).Top - 5, LEM_WIDTH + 8, LEM_HEIGHT + 5, picClear.HDC, 0, 0, vbSrcCopy)
        Else
clearLemming:
            Call BitBlt(picBoard.HDC, lemAll(Index).Left, lemAll(Index).Top, LEM_WIDTH, LEM_HEIGHT, picClear.HDC, 0, 0, vbSrcCopy)
        End If
    End If
End Sub

' ******************************
' This sub adds a lemming
' ******************************
Private Sub addLemming(Optional ByVal blnStart As Boolean = False)
    If UBound(lemAll()) >= intLemmings Then Exit Sub
    If blnStart = False Then ReDim Preserve lemAll(UBound(lemAll()) + 1) As Lemming
    
    lemAll(UBound(lemAll())).Alive = True
    lemAll(UBound(lemAll())).Left = intStartLeft
    lemAll(UBound(lemAll())).Top = intStartTop
    lemAll(UBound(lemAll())).PushX = LEM_STEP
    lemAll(UBound(lemAll())).FallFrame = LEM_STEP
    lemAll(UBound(lemAll())).Walking = False
    
    lblStatus.Caption = "Count: " & UBound(lemAll()) + 1
End Sub

' ******************************
' This function checks if a lemming
' is on ground or not. True = on ground
' ******************************
Private Function checkFeet(ByVal Index As Integer) As Boolean
    Dim i As Integer
    
    checkFeet = False
    For i = LEM_WIDTH To 0 Step -1
        If picBoard.Point(lemAll(Index).Left + LEM_WIDTH - i, lemAll(Index).Top + LEM_HEIGHT) <> RGB(0, 0, 0) Then
            checkFeet = True
            Exit For
        End If
    Next
    
    ' ******************************
    ' Kill lemming if off bottom of board
    ' ******************************
    If lemAll(Index).Top + LEM_HEIGHT = picBoard.ScaleHeight Then Call killLemming(Index)
End Function

Private Sub killLemming(ByVal Index As Integer, Optional ByVal blnExplode As Boolean = True, Optional ByVal strSound As String)
    lemAll(Index).Alive = False
    Call drawLemming(Index, True)
    
    If Not strSound = "" Then
        Call PlaySound(App.Path & "\sounds\" & strSound & ".wav")
        Else
        Call PlaySound(App.Path & "\sounds\die.wav")
    End If
    
    If blnExplode = True Then Call explodeLemming(Index)
End Sub

' ******************************
' Explode lemming nothing yet
' ******************************
Private Sub explodeLemming(ByVal Index As Integer)

End Sub
