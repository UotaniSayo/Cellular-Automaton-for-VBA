VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} status 
   Caption         =   "状态"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2175
   OleObjectBlob   =   "status.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnManual_Click()
    refresh
End Sub

Sub btnStart_Click()
    Dim shObj
    flag = Not flag
    Set shObj = CreateObject("WScript.Shell")
    If flag Then
        btnStart.Caption = "暂停"
        While flag
            refresh
            
        Wend
    Else
        btnStart.Caption = "继续"
    End If
    
End Sub

'刷新一次
Function refresh()
    Dim world2(1 To 302, 1 To 302) As Integer
    Dim count, life As Integer
    life = 0
    refreshCount = refreshCount + 1
    For i = 2 To 301
        For j = 2 To i
            world2(i, j) = world(i, j)
            world2(j, i) = world(j, i)
        Next j
    Next i
    
    For i = 2 To 301
        For j = 2 To 301
            count = world2(i - 1, j - 1) + world2(i, j - 1) + world2(i + 1, j - 1) + world2(i - 1, j) + world2(i + 1, j) + world2(i - 1, j + 1) + world2(i, j + 1) + world2(i + 1, j + 1)
            If count = 2 Then
                'If world(i, j) = 0 Then
                    'Range(xy2table((i), (j))).Interior.Color = 255
                'End If
                world(i, j) = 1
                life = life + 1
            Else
                'If world(i, j) = 1 Then
                    'Range(xy2table((i), (j))).Interior.Pattern = xlNone
                'End If
                world(i, j) = 0
            End If
        Next j
    Next i
    txtCount.Caption = "当前：" & life
    Range("A" & refreshCount).FormulaR1C1 = life
End Function
