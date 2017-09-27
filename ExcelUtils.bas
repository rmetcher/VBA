Attribute VB_Name = "ExcelUtils"
Option Explicit

Sub ExcelColumsToNumMap()

    '================================================================================
	' Function	: ExcelColumsToNumMap()
	' Args		: Null
	' Return	: Null
	' Output	: Debug.Print
    ' Author    : Rob Metcher
    ' Date      : 28/09/2017
	' Versions	: v1.0	Initial Commit
    ' Purpose   : Display excel columns in Alpha and Alphanumeric for easy translation
    '
	'================================================================================

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
