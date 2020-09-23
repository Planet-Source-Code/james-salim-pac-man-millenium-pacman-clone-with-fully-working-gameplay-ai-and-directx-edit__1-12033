VERSION 5.00
Begin VB.Form ItemProp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Item Properties"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1935
   Icon            =   "ItemProp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fItem 
      Caption         =   "&Item"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.TextBox txtItemAmount 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtItemDelay 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtItemAppearTime 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Frame fItemChoose 
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton oItemBeer 
            Caption         =   "Beer"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton oItemBerry 
            Caption         =   "Berry"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton oItemCherry 
            Caption         =   "Cherry"
            Height          =   375
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton oItemLife 
            Caption         =   "Life"
            Height          =   375
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Label lItemAmount 
         Caption         =   "Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lItemDelay 
         Caption         =   "Delay"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lAppearTime 
         Caption         =   "Appear Time"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   615
      End
   End
End
Attribute VB_Name = "ItemProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    oItemBeer = True
    RefreshItemProp
End Sub

Sub RefreshItemProp()
    Dim TempItem As LevelItemChar
    
    If oItemBeer Then lvledpropItemSelected = 1
    If oItemBerry Then lvledpropItemSelected = 2
    If oItemCherry Then lvledpropItemSelected = 3
    If oItemLife Then lvledpropItemSelected = 4
    
    With lvledBody.lvlItems
        Select Case lvledpropItemSelected
            Case 1: TempItem = .Beer
            Case 2: TempItem = .Berry
            Case 3: TempItem = .Cherry
            Case 4: TempItem = .Life
        End Select
    End With
            
    
    txtItemAppearTime = TempItem.AppearTime
    txtItemDelay = TempItem.Delay
    txtItemAmount = TempItem.Amount
End Sub

Sub SetItemProp(PropNo As Integer, Value As Integer)
    Dim TempItem As LevelItemChar
    With lvledBody.lvlItems
        Select Case lvledpropItemSelected
            Case 1: TempItem = .Beer
            Case 2: TempItem = .Berry
            Case 3: TempItem = .Cherry
            Case 4: TempItem = .Life
        End Select
    
    Select Case PropNo
        Case 1: TempItem.AppearTime = Value
        Case 2: TempItem.Delay = Value
        Case 3: TempItem.Amount = Value
    End Select

        Select Case lvledpropItemSelected
            Case 1: .Beer = TempItem
            Case 2: .Berry = TempItem
            Case 3: .Cherry = TempItem
            Case 4: .Life = TempItem
        End Select
    End With
End Sub

Private Sub oItemBeer_Click()
    RefreshItemProp
End Sub

Private Sub oItemBerry_Click()
    RefreshItemProp
End Sub

Private Sub oItemCherry_Click()
    RefreshItemProp
End Sub

Private Sub oItemLife_Click()
    RefreshItemProp
End Sub

Private Sub txtItemAmount_Change()
    SetItemProp 3, Val(txtItemAmount)
End Sub

Private Sub txtItemAppearTime_Change()
    SetItemProp 1, Val(txtItemAppearTime)
End Sub

Private Sub txtItemDelay_Change()
    SetItemProp 2, Val(txtItemDelay)
End Sub

