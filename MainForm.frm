VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DLL Base Address Generator"
   ClientHeight    =   600
   ClientLeft      =   10185
   ClientTop       =   5145
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   4770
   Begin VB.CommandButton generate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox dlladdr 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Address"
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'// 65536 * 256
'// DLL Base Addresses must be a multiple of 64K (65536)
'// LowerBound = 16777216 (which is 64K * 256)
'// UpperBound = 2147483648 (which is 64K * 32768)
'// 1) Generate Random Number between 256 and 32768
'// 2) Multiply by 64K
'// 3) Convert generated value to a Long value
'// 4) Convert generated value to Hex
'// 5) Convert generated string to lowercase (so that the letters are lower case)
'// 6) Add &H in front of string
Randomize
dlladdr = "&H" & LCase(Hex(CLng((32768 - 256 + 1) * Rnd + 256) * 65536))
End Sub

Private Sub generate_Click()
Randomize
dlladdr = "&H" & LCase(Hex(CLng((32768 - 256 + 1) * Rnd + 256) * 65536))
End Sub
