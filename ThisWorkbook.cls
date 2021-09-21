VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Helloword()
Dim Testmsg As String
Testmsg = "HelloUsers!"
MsgBox (Testmsg)
End Sub
Sub Hello()
 Dim Testmsg As String
 Testmsg = "Hello!"
 MsgBox (Testmsg)
End Sub
Sub DQ_analysis()
   Worksheets("DQ_Analysis").Activate
   Range("C1").Value = "50"
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
   
End Sub
