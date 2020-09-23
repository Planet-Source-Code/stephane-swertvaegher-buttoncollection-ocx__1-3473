VERSION 5.00
Object = "{4F1E23CB-AD57-11D2-B62A-00201853CFCC}#5.0#0"; "ButtonCollection.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "THEMA ButtonCollection Example"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2895
   Icon            =   "ButtonCollectionExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin ButtonCollection.UButton UButton1 
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   3900
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   344
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Caption         =   "Quit ButtonCollection Example"
   End
   Begin ButtonCollection.FadeButton FadeButton1 
      Height          =   495
      Left            =   420
      TabIndex        =   4
      Top             =   2760
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   873
      BackColor       =   13154464
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Picture         =   "ButtonCollectionExample.frx":0442
      FadePercent     =   80
      FadeOutValue    =   80
      FadeSpeed       =   5
   End
   Begin ButtonCollection.GrayButton GrayButton1 
      Height          =   495
      Left            =   420
      TabIndex        =   1
      Top             =   1740
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   13154464
      Picture         =   "ButtonCollectionExample.frx":0894
   End
   Begin ButtonCollection.BButton BButton1 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "About"
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
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
   Begin ButtonCollection.GrayButton GrayButton2 
      Height          =   495
      Left            =   1020
      TabIndex        =   2
      Top             =   1740
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   13154464
      Picture         =   "ButtonCollectionExample.frx":0CE6
   End
   Begin ButtonCollection.GrayButton GrayButton3 
      Height          =   495
      Left            =   1620
      TabIndex        =   3
      Top             =   1740
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColor       =   13154464
      Picture         =   "ButtonCollectionExample.frx":1138
   End
   Begin ButtonCollection.FadeButton FadeButton2 
      Height          =   495
      Left            =   1020
      TabIndex        =   5
      Top             =   2760
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   873
      BackColor       =   13154464
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Picture         =   "ButtonCollectionExample.frx":158A
      FadePercent     =   80
      FadeOutValue    =   80
      FadeSpeed       =   5
   End
   Begin ButtonCollection.FadeButton FadeButton3 
      Height          =   495
      Left            =   1620
      TabIndex        =   6
      Top             =   2760
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   873
      BackColor       =   13154464
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Picture         =   "ButtonCollectionExample.frx":19DC
      FadePercent     =   80
      FadeOutValue    =   80
      FadeSpeed       =   5
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BButton1_Click()
MsgBox "THEMA ButtonCollection by Maarten van Gompel, please visit the THEMA Homepage at: http://thema.cjb.net"
End Sub


Private Sub UButton1_Click()
End
End Sub
