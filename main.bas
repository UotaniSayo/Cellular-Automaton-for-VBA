Attribute VB_Name = "模块1"

Public flag As Boolean
Public world(1 To 302, 1 To 302) As Integer
Public refreshCount As Integer

Public Sub cell()
    'table init
    'Rows("1:102").RowHeight = 5
    'Columns("A:CX").ColumnWidth = 0.5

    'Range(xy2table(40, 41)).Interior.Color = 255
    'Range(xy2table(41, 41)).Interior.Color = 255
    'Range(xy2table(41, 42)).Interior.Color = 255
    
    'var init
    flag = False
    refreshCount = 0
    
    For i = 1 To 302
        For j = 1 To i
            world(i, j) = 0
            world(j, i) = 0
        Next j
    Next i
    world(140, 141) = 1
    world(141, 141) = 1
    world(141, 142) = 1
    
    'load window
    Load status
    status.Show
End Sub


Public Function xy2table(x As Integer, y As Integer) As String
    Dim letter As Variant
    Dim out As String
    letter = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    If x <= 26 Then
        out = letter(x - 1)
    'perhaps x mod 27 will be better?
    ElseIf x Mod 26 = 0 Then
        out = letter(x \ 26 - 2) & "Z"
    Else
        out = letter(x \ 26 - 1) & letter(x Mod 26 - 1)
    End If
    xy2table = out & y
End Function


