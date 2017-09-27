Attribute VB_Name = "ExcelUtils"
Option Explicit

Sub ExcelColumsToNumMap()
    Dim ColNo   As Long
    Dim ColTemp As Long
    
    ColNo = 1
    ColTemp = 1
    
    Do While ColNo < 234
        For ColTemp = ColNo To (ColNo + 25)
            Debug.Print Split(Cells(, ColTemp).Address, "$")(1) & vbTab;
        Next
        
        Debug.Print ""
        
        For ColTemp = ColNo To (ColNo + 25)
            Debug.Print ColTemp & vbTab;
        Next
        
        Debug.Print ""
        Debug.Print ""
        
        ColNo = ColTemp
    Loop
End Sub
