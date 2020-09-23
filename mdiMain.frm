VERSION 5.00
Object = "{A0BCF497-C849-4D13-B171-DCA141B37DD8}#2.0#0"; "SplitBox10.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin SplitBox10.SplitBox SplitBox2 
      Align           =   3  'Align Left
      Height          =   8265
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Split Box 2"
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   14579
      Begin VB.PictureBox Picture2 
         Height          =   8115
         Left            =   30
         ScaleHeight     =   8055
         ScaleWidth      =   3795
         TabIndex        =   2
         Top             =   90
         Width           =   3855
      End
   End
   Begin SplitBox10.SplitBox SplitBox1 
      Align           =   4  'Align Right
      Height          =   8265
      Left            =   6210
      TabIndex        =   0
      ToolTipText     =   "Split Box 1"
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   14579
      Begin VB.PictureBox Picture1 
         Height          =   8130
         Left            =   90
         ScaleHeight     =   8070
         ScaleWidth      =   3810
         TabIndex        =   3
         Top             =   75
         Width           =   3870
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SplitBox2_Resize()

Picture2.Width = SplitBox2.Width - 150
Picture2.Height = SplitBox2.Height - 150

End Sub

Private Sub splitbox1_resize()

Picture1.Width = SplitBox1.Width - 150
Picture1.Height = SplitBox1.Height - 150

End Sub
