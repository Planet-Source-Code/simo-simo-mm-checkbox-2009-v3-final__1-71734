VERSION 5.00
Object = "*\A..\mm_CheckBox.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00FFDBBF&
   Caption         =   "MM CheckBox 2009 (v3)"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Désable All"
      Height          =   315
      Left            =   6870
      TabIndex        =   13
      Top             =   1935
      Width           =   1200
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Check All"
      Height          =   315
      Left            =   5610
      TabIndex        =   0
      Top             =   1935
      Value           =   1  'Checked
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   2385
      Left            =   3240
      TabIndex        =   16
      Top             =   -60
      Width           =   7440
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   17
         Top             =   450
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   423
         Checked         =   -1  'True
         RoundedValue    =   8
         Activecolor     =   8388608
         Caption         =   "Syntaxe Vérification"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   8388608
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   6
         Left            =   195
         TabIndex        =   18
         Top             =   810
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   423
         Checked         =   -1  'True
         RoundedValue    =   8
         Activecolor     =   8421376
         desActivecolor  =   0
         Caption         =   "Compilation in Background"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   4210688
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   7
         Left            =   195
         TabIndex        =   19
         Top             =   1140
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   423
         Checked         =   -1  'True
         RoundedValue    =   8
         Activecolor     =   33023
         desActivecolor  =   0
         Caption         =   "Activer the débogage"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   33023
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   8
         Left            =   195
         TabIndex        =   20
         Top             =   1560
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   423
         Enabled         =   0   'False
         Checked         =   -1  'True
         RoundedValue    =   9
         Activecolor     =   16384
         desActivecolor  =   0
         Caption         =   "Use Key"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   16384
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   9
         Left            =   4230
         TabIndex        =   21
         Top             =   450
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   423
         Checked         =   -1  'True
         RoundedValue    =   9
         Activecolor     =   0
         desActivecolor  =   0
         Caption         =   "Syntaxe Vérification         "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   10
         Left            =   4230
         TabIndex        =   22
         Top             =   810
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   423
         Checked         =   -1  'True
         RoundedValue    =   9
         Activecolor     =   192
         desActivecolor  =   0
         Caption         =   "Compilation in Background"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   192
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   11
         Left            =   4230
         TabIndex        =   23
         Top             =   1140
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   423
         Checked         =   -1  'True
         RoundedValue    =   9
         Activecolor     =   12583104
         desActivecolor  =   0
         Caption         =   "Activer the débogage       "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   12583104
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   12
         Left            =   4230
         TabIndex        =   24
         Top             =   1560
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   423
         Enabled         =   0   'False
         Checked         =   -1  'True
         RoundedValue    =   9
         Activecolor     =   0
         desActivecolor  =   0
         Caption         =   "Use Key                             "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   13
         Left            =   3435
         TabIndex        =   25
         Top             =   450
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   423
         Checked         =   -1  'True
         RoundedValue    =   10
         Activecolor     =   255
         desActivecolor  =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   14
         Left            =   3420
         TabIndex        =   26
         Top             =   810
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   423
         Checked         =   -1  'True
         RoundedValue    =   10
         Activecolor     =   32896
         desActivecolor  =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   15
         Left            =   3420
         TabIndex        =   27
         Top             =   1140
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   423
         Checked         =   -1  'True
         RoundedValue    =   10
         Activecolor     =   16711680
         desActivecolor  =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox2 
         Height          =   240
         Index           =   16
         Left            =   3420
         TabIndex        =   28
         Top             =   1560
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   423
         Enabled         =   0   'False
         Checked         =   -1  'True
         RoundedValue    =   10
         Activecolor     =   8421631
         desActivecolor  =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " Without Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   2940
         TabIndex        =   31
         Top             =   90
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " With Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   4935
         TabIndex        =   30
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " With Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   795
         TabIndex        =   29
         Top             =   135
         Width           =   1260
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   3975
         X2              =   3975
         Y1              =   285
         Y2              =   1830
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   3180
         X2              =   3180
         Y1              =   285
         Y2              =   1830
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Checked ?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1140
      TabIndex        =   1
      Top             =   1275
      Width           =   990
   End
   Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox1 
      Height          =   450
      Left            =   1290
      TabIndex        =   2
      Top             =   540
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   794
      Checked         =   -1  'True
      Small           =   0   'False
      RoundedValue    =   25
      Activecolor     =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2265
      Left            =   -165
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2265
      ScaleWidth      =   3360
      TabIndex        =   3
      Top             =   30
      Width           =   3360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFDBBF&
      Height          =   2625
      Left            =   45
      TabIndex        =   4
      Top             =   2265
      Width           =   10635
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   1020
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   12582912
         Caption         =   "Test Number 1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   12582912
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   2
         Left            =   9150
         TabIndex        =   6
         Top             =   990
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   3
         Left            =   9150
         TabIndex        =   7
         Top             =   1980
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Enabled         =   0   'False
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   12583104
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   5
         Left            =   135
         TabIndex        =   8
         Top             =   1515
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Caption         =   "Test Caption (a different Color for Caption)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   16576
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   6
         Left            =   135
         TabIndex        =   9
         Top             =   510
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   8421376
         Caption         =   "This one is : On"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   8421376
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   7
         Left            =   9150
         TabIndex        =   10
         Top             =   1455
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   49152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   4
         Left            =   135
         TabIndex        =   11
         Top             =   2025
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   794
         Enabled         =   0   'False
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   33023
         Caption         =   "&Juste a simple Caption"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   33023
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   1
         Left            =   9150
         TabIndex        =   12
         Top             =   495
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   8
         Left            =   4485
         TabIndex        =   32
         Top             =   1020
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   32896
         Caption         =   "Teste Number 1                             "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   16448
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   9
         Left            =   4485
         TabIndex        =   33
         Top             =   1515
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   128
         Caption         =   "Test Caption (a different Color for Caption)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   192
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   10
         Left            =   4485
         TabIndex        =   34
         Top             =   510
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   794
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   0
         Caption         =   "Hi everyone                                   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   33023
      End
      Begin MM_Adv_CheckBox.mm_checkbox mm_checkbox3 
         Height          =   450
         Index           =   11
         Left            =   4485
         TabIndex        =   35
         Top             =   2025
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   794
         Enabled         =   0   'False
         Checked         =   -1  'True
         Small           =   0   'False
         RoundedValue    =   30
         Activecolor     =   33023
         Caption         =   "&Juste a simple Caption                              "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColor    =   33023
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   " With Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   1
         Left            =   30
         TabIndex        =   15
         Top             =   135
         Width           =   8775
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   8835
         X2              =   8835
         Y1              =   180
         Y2              =   2550
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   " Without Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   0
         Left            =   8880
         TabIndex        =   14
         Top             =   135
         Width           =   1710
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFDBBF&
      Height          =   1110
      Left            =   45
      TabIndex        =   36
      Top             =   4845
      Width           =   10635
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "White"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFDBBF&
         Caption         =   "blue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Checked true/false"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4980
         TabIndex        =   38
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Enabled true/false"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3705
         TabIndex        =   37
         Top             =   300
         Width           =   1275
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   3285
         X2              =   3285
         Y1              =   255
         Y2              =   870
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()

On Error Resume Next

For i = 0 To 20
    Me.mm_checkbox2(i).Checked = Check1.Value
Next i

End Sub



Private Sub Check3_Click()
On Error Resume Next

For i = 0 To 20
    Me.mm_checkbox2(i).Enabled = IIf(Check3.Value = 1, False, True)
Next i

End Sub

Private Sub Command1_Click()

MsgBox "Button Statut : " & mm_checkbox1.Checked

End Sub

Private Sub Command2_Click()
'On Error Resume Next
mm_checkbox3(0).Checked = Not mm_checkbox3(0).Checked
For i = 1 To 11
    mm_checkbox3(i).Checked = mm_checkbox3(0).Checked = True ' False 'Not mm_checkbox3(i).Checked
Next i

End Sub

Private Sub Command3_Click()
'On Error Resume Next
mm_checkbox3(0).Enabled = Not mm_checkbox3(0).Enabled
For i = 1 To 11
    mm_checkbox3(i).Enabled = mm_checkbox3(0).Enabled
Next i

End Sub


Private Sub Command4_Click()
Frame1.BackColor = vbWhite
'On Error Resume Next
For i = 0 To 11 '20 '7
    mm_checkbox3(i).RoundedValue = 23
Next i


End Sub

Private Sub Command5_Click()
Frame1.BackColor = &HFFDBBF
'On Error Resume Next
For i = 0 To 11 '20 '7
    mm_checkbox3(i).RoundedValue = 27
Next i


End Sub

Private Sub Form_Load()

On Error Resume Next

For i = 0 To 20

    Me.mm_checkbox2(i).desActivecolor = &H808080   'RGB(r, g, b)

Next i


For i = 0 To 9
    Me.mm_checkbox3(i).RoundedValue = 31
Next i

End Sub



Private Sub HScroll1_Change()
HScroll1_Scroll
End Sub

Private Sub mm_checkbox3_Click(Index As Integer)
If Index = 6 Then
    mm_checkbox3(6).Caption = "This one is : " & IIf(mm_checkbox3(6).Checked = True, "On", "Off")
End If
End Sub


