Attribute VB_Name = "Module1"
Public lstindex As Long
Public Function newsong()
If Not lstindex = Form1.List1.ListCount - 1 Then
Form1.MediaPlayer1.FileName = Form1.List2.List(lstindex + 1)
Form1.List1.ListIndex = lstindex + 1
lstindex = lstindex + 1
End If
End Function
