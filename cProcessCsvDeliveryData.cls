VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProcessCsvDeliveryData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iFileNum    As Integer
Private oError      As ErrObject
Private oXlApp      As Excel.Application
Private oWb             As Workbook
Private oWs             As Worksheet
Private oCollDtz        As Collection

Private from_timezone   As String
Private to_timezone     As String

Public Property Get get_from_tz() As String
    get_from_tz = from_timezone
End Property

Public Property Let set_from_tz(tz As String)
    from_timezone = tz
End Property

Public Property Get get_to_tz() As String
    get_to_tz = to_timezone
End Property

Public Property Let set_to_tz(tz As String)
    to_timezone = tz
End Property

Private Sub Class_Initialize()
    from_timezone = "Africa/Johannesburg"
    to_timezone = "UTC"
End Sub
Public Function GetLastError() As ErrObject
    Dim a As Variant
    
    If Not oError Is Nothing Then
        Set GetLastError = oError
    Else
        Set GetLastError = Nothing
    End If

End Function

Public Function ETL_Execute() As Integer
    
    '==============================================================================
    ' Class     : cProcessCsvDeliveryData
    ' Function  : ETL_Execute
    ' Author    : Rob Metcher
    ' Date      : 19/09/2017
    ' Purpose   : Iterate through source folder .csv file to import as per mappings
    '
    '
    ' input csv         Z:\Engen\INPUTCSV\IN-DeliveryRecords\Robs_Test_Csv
    ' output datafile   \\epsserver.lob.local\EPS\AA_PRoductionTrials\BP Turkey\Vault8\BPT_MAHMUTBEYM\BP_TurkeyDeliveries_MAHMUTBEYM.xls
    ' log messages      \\epsserver.lob.local\EPS\AA_PRoductionTrials\Engen\ErrorTrap
    '
    ' SHEET MAPPING TO CSVFILE
    ' sheet('TerminalDeliveries').column(15) (gcBP_TurkeyRawDataTerminalDeliveries_LOBFromDTSCol) = invday + 00:00:01 to UTC
    ' sheet('TerminalDeliveries').column(16) (gcBP_TurkeyRawDataTerminalDeliveries_LOBToDTSCol) = invday + 23:59:59 to UTC
    ' sheet('TerminalDeliveries').column(8) (gcBP_TurkeyRawDataTerminalDeliveries_site_idCol) = siteid
    ' sheet('TerminalDeliveries').column(9) (gcBP_TurkeyRawDataTerminalDeliveries_site_codeCol) = site
    ' sheet('TerminalDeliveries').column(17) (gcBP_TurkeyRawDataTerminalDeliveries_LOB_ProdCodeCol) = 'prodid
    ' sheet('TerminalDeliveries').column(13) (gcBP_TurkeyRawDataTerminalDeliveries_product_codeCol) = prodshort
    ' xxxx = ufiname (Not used)
    ' sheet('TerminalDeliveries').column(18) (gcBP_TurkeyRawDataTerminalDeliveries_LOB_VOlumeL15Col) = deliv if L15flag = yes
    ' sheet('TerminalDeliveries').column(19) (gcBP_TurkeyRawDataTerminalDeliveries_LOB_VOlumeAmbCol) = deliv if L15flag = else
    '
    '===============================================================================
    
    Const message_log_path = "\\epsserver.lob.local\EPS\AA_PRoductionTrials\Engen\ErrorTrap\_DAEPT_V_66_msg.log"
    Const procName = "import_terminal_deliveries"
    Const module_name = "Utilities2"
    Const status_bar = True
    
    
    Dim sSourcePath     As String
    Dim sExtension      As String
    Dim sDestPath       As String
    Dim sDestFile       As String
    Dim sExcelPath      As String
    Dim sExcelFile      As String
    Dim sFile()         As String
    Dim sFileContent    As String
    Dim sLineArray()    As String
    Dim sSiteName       As String
    Dim sTemp()         As String
    Dim sFnameParts()   As String
    Dim sTempFile       As String
    Dim sErrMsg         As String
    Dim sStatusBarMsg   As String
    
    Dim vDataArray()    As Variant
    Dim vDateTimeTz     As Variant
    Dim vItem           As Variant
    
    
    Dim iRow            As Integer
    Dim iIndexOfStrings As Integer: iIndexOfStrings = 0
    Dim iStart          As Integer: iStart = 0
    Dim iStop           As Integer: iStop = 0
    Dim a, b            As Integer
    Dim iCsvFileIndex   As Integer
    Dim i               As Integer: i = 0
    Dim iFilesFound  As Integer: iFilesFound = 0
    
    Dim iIndex          As Long
    
'    Dim oXlApp          As Excel.Application
'    Dim oWb             As Workbook
'    Dim oWs             As Worksheet
'    Dim oCollDtz        As Collection
    
    On Error GoTo error_handler
    
    'err.Raise (6)
    
    '======================
    'Setup file definitions
    '======================
    sSourcePath = "Z:\Engen\INPUTCSV\IN-DeliveryRecords\Robs_Test_Csv\"
    sExtension = "*.csv" 'Target File Extension (must include wildcard "*")
    sDestPath = "\\epsserver.lob.local\EPS\AA_PRoductionTrials\Engen\Vault2\ENGEN_%1\" 'the site name is place in the %1 position
    sDestFile = "EngenDeliveries_%1.xls" 'the site name is place in the %1 position
    
    '====================
    'Status bar init
    '====================
    If (status_bar) Then
        Application.DisplayStatusBar = True
        sStatusBarMsg = "Csv Reader..."
        Application.StatusBar = sStatusBarMsg
    End If
    
    '===================================================
    'Create an instance of a specific version of excel
    '===================================================
    Set oXlApp = CreateObject("Excel.Application.11")
    
'    Excel 97 = 8
'    Excel 2000 = 9
'    Excel 2002 = 10
'    Excel 2003 = 11
'    Excel 2007 = 12
'    Excel 2010 = 14
'    Excel 2013 = 15
'    Excel 2016 = 16

    '===========================================
    ' populate a array full of .csv files found
    '===========================================
    If (status_bar) Then
        Application.StatusBar = sStatusBarMsg & " get csv files from " & sSourcePath
    End If
    
    iCsvFileIndex = 0
    ReDim sFile(iCsvFileIndex)
    sFile(iCsvFileIndex) = Dir(sSourcePath & sExtension)
    Do While sFile(iCsvFileIndex) <> ""
        iCsvFileIndex = iCsvFileIndex + 1
        ReDim Preserve sFile(iCsvFileIndex)
        sFile(iCsvFileIndex) = Dir
    Loop
    
    '==========================================
    'Loop through each .csv file in sFile array
    '==========================================
    For iCsvFileIndex = 0 To UBound(sFile) - 1
        
        'debug.print "Processing " & sFile(iCsvFileIndex) & " ";
        
        'Open the CSV file
        iFileNum = FreeFile
        Open sSourcePath & sFile(iCsvFileIndex) For Input As #iFileNum
        
        'Get the sitename part from the csv filename
        sFnameParts = Split(sFile(iCsvFileIndex), "_")
        sSiteName = sFnameParts(1)
        
        '=========================================
        'Populate an array with the files contents
        '=========================================
        
        i = 0
        
        ReDim vDataArray(i)
        Do While Not EOF(iFileNum) ' Loop until end of file.
            ReDim Preserve vDataArray(i) ' Redim the array for the new element
            Line Input #iFileNum, vDataArray(i) ' read next line from file and add text to the array
            i = i + 1
        Loop
                
        Close #iFileNum
        
        '============================================
        'Open appropriate excel workbook by site name
        '============================================
        sExcelFile = sDestFile
        sExcelPath = sDestPath
        sExcelFile = Replace(sDestFile, "%1", sSiteName)
        sExcelPath = Replace(sDestPath, "%1", sSiteName)
        
        If (status_bar) Then
            Application.StatusBar = sStatusBarMsg & sFile(iCsvFileIndex) & " to " & sExcelFile
        End If
        
        '======================================
        'check if the excel file/folder exists
        '======================================
        If Len(Dir(sExcelPath & sExcelFile)) = 0 Then
            message_log message_log_path, module_name, procName, , "File not found! " & sExcelPath & sExcelFile
            
            If (status_bar) Then
                Application.StatusBar = sStatusBarMsg & sFile(iCsvFileIndex) & " to " & sExcelFile & " (Not found)"
            End If
            
            'debug.print "File not found! " & sExcelPath & sExcelFile
            GoTo next_file
        Else
            'debug.print "File found! " & sExcelFile
            iFilesFound = iFilesFound + 1
            message_log message_log_path, module_name, procName, , "File found! " & sExcelPath & sExcelFile
            If (status_bar) Then
                Application.StatusBar = sStatusBarMsg & sFile(iCsvFileIndex) & " to " & sExcelFile & " (Found)"
            End If
        End If
        
        'debug.print ""
        
        oXlApp.Visible = False
        oXlApp.EnableEvents = False
        
        Set oWb = oXlApp.Workbooks.Open(FileName:=sExcelPath & sExcelFile, IgnoreReadOnlyRecommended:=True)
        'Set oWb = Workbooks.Open(FileName:=sExcelPath & sExcelFile)
        Set oWs = oWb.Worksheets("TerminalDeliveries")
        
        'Get lastused
'        iRow = oWs.Cells(oWs.Rows.Count, "H").End(xlUp).Row + 1
        
        iRow = oWs.Range("H" & oWs.Rows.Count).End(xlUp).Row + 1
        
'        With oWs.UsedRange
'            iRow = .Rows(.Rows.Count).Row + 1
'        End With
                
        ReDim sLineArray(iIndexOfStrings)
        For iIndex = LBound(vDataArray) + 1 To UBound(vDataArray)
            
            'Pull out all the strings between double quotes, ignoring comma delimiters
            sLineArray = get_substrings_from_string((vDataArray(iIndex)))
            
            'write the sLineArray data to the target excel data sheet
            
            vDateTimeTz = convtz(DateValue(sLineArray(0)) + TimeValue("00:00:01"), from_timezone, to_timezone) 'Python timezone call
            oWs.Cells(iRow, gcBP_TurkeyRawDataTerminalDeliveries_LOBFromDTSCol).Value = vDateTimeTz(0, 0)
            
            vDateTimeTz = convtz(DateValue(sLineArray(0)) + TimeValue("23:59:59"), from_timezone, to_timezone) 'Python timezone call
            oWs.Cells(iRow, gcBP_TurkeyRawDataTerminalDeliveries_LOBToDTSCol).Value = vDateTimeTz(0, 0)
            
            oWs.Cells(iRow, gcBP_TurkeyRawDataTerminalDeliveries_site_idCol).Value = sLineArray(1)
            oWs.Cells(iRow, gcBP_TurkeyRawDataTerminalDeliveries_site_codeCol).Value = sLineArray(2)
            oWs.Cells(iRow, gcBP_TurkeyRawDataTerminalDeliveries_LOB_ProdCodeCol).Value = sLineArray(3)
            oWs.Cells(iRow, gcBP_TurkeyRawDataTerminalDeliveries_product_codeCol).Value = sLineArray(4)
            
            
            'L15Flag
            If sLineArray(7) = UCase("YES") Then
                oWs.Cells(iRow, gcBP_TurkeyRawDataTerminalDeliveries_LOB_VOlumeL15Col).Value = sLineArray(6)
            Else
                oWs.Cells(iRow, gcBP_TurkeyRawDataTerminalDeliveries_LOB_VOlumeAmbCol).Value = sLineArray(6)
            End If
            
            iRow = iRow + 1
        Next iIndex
        
        'oXlApp.DisplayAlerts = False
        
        oWb.CheckCompatibility = False
        
        If oWb.Saved = False Then
            oWb.Save
        End If
        
        oWb.Close
        
        Set oWs = Nothing
        Set oWb = Nothing
        'Set oXlApp = Nothing
'Stop
next_file:
    'Loop
    Next
    
    'oWb.Close SaveChanges:=False
'    Set oWs = Nothing
'    Set oWb = Nothing
'
'    Set oXlApp = Nothing
    
    message_log message_log_path, module_name, procName, , "Done :=)"
    
    If (status_bar) Then
        If iFilesFound <> UBound(sFile) Then
            Application.StatusBar = sStatusBarMsg & "Done! " & iFilesFound & " Excel files of " & UBound(sFile) & " Csv files found. See message log at " & message_log_path
        Else
            Application.StatusBar = sStatusBarMsg & "Done! " & iFilesFound & " Excel files of " & UBound(sFile) & " Csv files found!"
        End If
    End If
    
    ETL_Execute = 1
    
    Exit Function
    
error_handler:
    
    Set oError = err
    
    sErrMsg = "Error number: " & Str(err.Number) & " Source: " & err.Source & " Description: " & err.Description
           
    message_log message_log_path, module_name, procName, , sErrMsg
    
'    If Not oWs Is Nothing Then
'        Set oWs = Nothing
'    End If
'
'    If Not oWb Is Nothing Then
'        oWb.CheckCompatibility = False
'
'        If oWb.Saved = False Then
'            oWb.Save
'        End If
'
'        oWb.Close
'        Set oWb = Nothing
'    End If
'
'    If Not oXlApp Is Nothing Then
'        Set oXlApp = Nothing
'    End If
'
'    Close #iFileNum
    
    If (status_bar) Then
        Application.StatusBar = sStatusBarMsg & "Errors - see log (" & message_log_path & ")"
    End If
    
    ETL_Execute = 0
    
End Function

Private Function get_substrings_from_string(sInString As String) As String()
    Dim sLineArray()    As String
    
    Dim iIndexOfStrings As Integer
    Dim a, b            As Integer
    Dim iStart          As Integer
    Dim iStop           As Integer
    Dim iIndex          As Integer
    
    iStart = 0
    iStop = 0
    iIndexOfStrings = 0
    ReDim sLineArray(iIndexOfStrings)
    For a = 0 To 7
        ReDim Preserve sLineArray(iIndexOfStrings)
        
        'get position of double quotes
        iStart = InStr(iStart + 1, sInString, """")
        iStop = InStr(iStart + 1, sInString, """")
        
        'read in string between double quotes
        For b = iStart + 1 To iStop - 1
            sLineArray(iIndexOfStrings) = sLineArray(iIndexOfStrings) & Mid(sInString, b, 1)
        Next
        iStart = iStop
        iIndexOfStrings = iIndexOfStrings + 1
    Next

    get_substrings_from_string = sLineArray
End Function
Private Sub message_log(sLogPath As String, module_name As String, procName As String, Optional err As Variant = Nothing, Optional comments As String = "")
    Dim iFileNumErr As Integer
    Dim sOutput     As String
    
    sOutput = Format(Now(), "yyyy-mm-dd hh:mm") & " - Log: " & module_name & "." & procName
    
    If (Not err Is Nothing) Then
        sOutput = sOutput & ": Error# " & err.Number & " """ & err.Description & """"
    End If
    
    If Len(comments) > 0 Then
        sOutput = sOutput & " : " & comments
    End If
    
    iFileNumErr = FreeFile
    Open sLogPath For Append As #iFileNumErr
    Print #iFileNumErr, sOutput
    Close #iFileNumErr
    
End Sub

Private Sub Class_Terminate()
    
    If Not oWs Is Nothing Then
        Set oWs = Nothing
    End If
    
    If Not oWb Is Nothing Then
        oWb.CheckCompatibility = False
        
        If oWb.Saved = False Then
            oWb.Save
        End If
        
        oWb.Close
        Set oWb = Nothing
    End If
    
    If Not oXlApp Is Nothing Then
        Set oXlApp = Nothing
    End If
    
    Close #iFileNum
    
End Sub
