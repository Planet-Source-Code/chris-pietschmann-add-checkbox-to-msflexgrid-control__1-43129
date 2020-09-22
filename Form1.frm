VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "MSFlexGrid Checkbox Example"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4260
      _Version        =   393216
      AllowBigSelection=   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Chris Pietschmann
'http://PietschSoft.itgo.com

Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub Form_Load()
        With MSFlexGrid1
            .Rows = 10
            .Cols = 3
            
            .AllowUserResizing = flexResizeBoth
                        
            'name the cols
            For i = 1 To .Cols - 1
                .Row = 0
                .Col = i
                .Text = "Column " & i
            Next i
            
            'name the rows
            For i = 1 To .Rows - 1
                .Col = 0
                .Row = i
                .Text = "Row " & i
            Next i
                        
            'define fields as checkbox
            For y = 1 To .Rows - 1
                For x = 1 To .Cols - 1
                    .Row = y
                    .Col = x
                    .CellFontName = "Wingdings"
                    .CellFontSize = 14
                    .CellAlignment = flexAlignCenterCenter
                    .Text = strUnChecked
                Next x
            Next y
        End With
    
End Sub

Private Sub Form_Resize()
    MSFlexGrid1.Width = Me.ScaleWidth
    MSFlexGrid1.Height = Me.ScaleHeight
End Sub

Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
        With MSFlexGrid1
            If .TextMatrix(iRow, iCol) = strUnChecked Then
                .TextMatrix(iRow, iCol) = strChecked
            Else
                .TextMatrix(iRow, iCol) = strUnChecked
            End If
        End With
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
        With MSFlexGrid1
            Call TriggerCheckbox(.Row, .Col)
        End With
    End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With MSFlexGrid1
            If .MouseRow <> 0 And .MouseCol <> 0 Then
                Call TriggerCheckbox(.MouseRow, .MouseCol)
            End If
        End With
    End If
End Sub
