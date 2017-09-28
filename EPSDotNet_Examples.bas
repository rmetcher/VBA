Attribute VB_Name = "EPSDotNet_Examples"
Option Explicit

Dim controller As EpsDotNet.FuelSystemController
Dim siteID As Integer


Private Sub test()

    Dim FuelSystemControllerFactory As EpsDotNet.FuelSystemControllerFactory
    
    Set FuelSystemControllerFactory = New EpsDotNet.FuelSystemControllerFactory
    Set controller = FuelSystemControllerFactory.CreateFuelSystemController("epsdb-staging-cluster-1.cluster-cuvifi2nhzwe.us-east-2.rds.amazonaws.com", "EPS", "lob", "ss_password")
    
    siteID = 2
    
    AddDispenserWithPumpHoseHosePortion controller
    AddNewFuelSystemWithGradeTankHosePortion controller
    GetActiveFuelSystem controller
End Sub

'Variant Count
Function variantCount(variantObj As Variant)
    variantCount = UBound(variantObj) - LBound(variantObj) + 1
End Function

company

Function GetCompany(controller As EpsDotNet.FuelSystemController)
    Dim companyCode As String
    Dim company As EpsDotNet.company
    
    companyCode = "AW0001"
    Set company = controller.GetCompany(companyCode)
   
End Function

'site
Function GetSites(controller As EpsDotNet.FuelSystemController)

    Dim companyID As Long
    companyID = 3
    Dim sites As Variant
    sites = controller.GetSites(companyID)
    Dim sitesCount As Integer
    
    For sitesCount = 0 To UBound(sites)
   
        Dim site As EpsDotNet.site
        Set site = sites(sitesCount)
        MsgBox "Site Name " + site.siteName
        
   Next sitesCount
   MsgBox "Site Count for Company ID " + CStr(companyID) + " is " + CStr(sitesCount)

End Function

'tank
' Need to know valid  SiteID  from DB
Function AddTank(controller As EpsDotNet.FuelSystemController)

    Dim tank As EpsDotNet.tank
    Set tank = New EpsDotNet.tank
    tank.siteID = 33
    tank.ForecourtTankNo = 5
    
    Dim returnTank As EpsDotNet.tank
    Set returnTank = controller.CommitChangesTanK(tank)

End Function

Function GetTanks(controller As EpsDotNet.FuelSystemController)

    'Dim siteID As Long
    'siteID = 33
    Dim tanks As Variant
    
    tanks = controller.GetTanks(siteID)
    Dim tanksCount As Integer
    
    For tanksCount = 0 To UBound(tanks)
   
        Dim tank As EpsDotNet.tank
        Set tank = tanks(tanksCount)
        MsgBox "Tank No " + CStr(tank.ForecourtTankNo)
        
   Next tanksCount
   MsgBox "Site " + CStr(siteID) + " Contains " + CStr(tanksCount) + " tanks"
   
End Function

'grade
Function AddNewGradeByName(controller As EpsDotNet.FuelSystemController)

    Dim gradeName As String
    gradeName = "Grade1"
    
    Dim companyID As Integer
    companyID = 3
    
    'Dim siteID As Integer
    'siteID = 33
    
    Dim grade As EpsDotNet.grade
    Set grade = controller.AddNewGrade(companyID, siteID, gradeName)
    
End Function

Function GetGrades(controller As EpsDotNet.FuelSystemController)

    Dim companyID As Long
    companyID = 3
    Dim grades As Variant
    grades = controller.GetGrades(companyID)
    Dim gradesCount As Integer
    
    For gradesCount = 0 To UBound(grades)
   
        Dim grade As EpsDotNet.grade
        Set grade = grades(gradesCount)
        MsgBox "Grade Name " + grade.gradeName
        
   Next gradesCount
   MsgBox "Grade Count for Company ID " + CStr(companyID) + " is " + CStr(gradesCount)
   
End Function

'dispenser
Function AddDispenser(controller As EpsDotNet.FuelSystemController)

    Dim hose1 As EpsDotNet.hose
    Set hose1 = New EpsDotNet.hose
    hose1.ForecourtHoseNumber = 1
    
    Dim hose2 As EpsDotNet.hose
    Set hose2 = New EpsDotNet.hose
    hose2.ForecourtHoseNumber = 2

    Dim pump As EpsDotNet.pump
    Set pump = New EpsDotNet.pump
    pump.ForecourtPumpNumber = 1
    pump.AddHose hose1
    pump.AddHose hose2

    Dim dispenser As EpsDotNet.dispenser
    Set dispenser = New EpsDotNet.dispenser
    dispenser.DispenserName = "Dispenser Name"
    dispenser.siteID = 33
    dispenser.SerialNumber = "SER 1"
    dispenser.DispenserNumber = 1
    dispenser.AddPump pump

    Dim returnDispenser As EpsDotNet.dispenser
    Set returnDispenser = controller.CommitChangesDispenser(dispenser)

End Function

' Need to know valid  SiteID and Grade ID from DB
Function AddDispenserWithPumpHoseHosePortion(controller As EpsDotNet.FuelSystemController)

    Dim hose1 As EpsDotNet.hose
    Set hose1 = New EpsDotNet.hose
    hose1.ForecourtHoseNumber = 7
    
    Dim hose2 As EpsDotNet.hose
    Set hose2 = New EpsDotNet.hose
    hose2.ForecourtHoseNumber = 8
    
    Dim hosePortion As EpsDotNet.hosePortion
    Set hosePortion = New EpsDotNet.hosePortion
    
    Dim companyID As Long
    companyID = 1
    Dim gradesVariant As Variant
    gradesVariant = controller.GetGrades(companyID)
    
    ' Only add if valid grade is available
    If (variantCount(gradesVariant) > 0) Then
        Dim grade As grade
        Set grade = gradesVariant(0)
        hosePortion.GradeId = grade.GradeId
        hose1.AddHosePortion hosePortion
        
        Dim pump As EpsDotNet.pump
        Set pump = New EpsDotNet.pump
        pump.ForecourtPumpNumber = 1
        pump.AddHose hose1
        pump.AddHose hose2
        
        Dim dispenser As EpsDotNet.dispenser
        Set dispenser = New EpsDotNet.dispenser
        
        dispenser.DispenserName = "Dispenser Name"
        'dispenser.siteID = 31
        dispenser.siteID = siteID
        dispenser.SerialNumber = "SER 1"
        dispenser.DispenserNumber = 4
        dispenser.AddPump pump
        
        controller.CommitChangesDispenser dispenser
        
    End If
End Function

'Load all dispensers and hose info for a site
Function GetDispensers(controller As EpsDotNet.FuelSystemController)

 'Dim siteID As Long
 'siteID = 33
 
 Dim dispensers As Variant
 dispensers = controller.GetDispensers(siteID)

 Dim dispensersCount As Integer
 For dispensersCount = 0 To UBound(dispensers)
   Dim dispenser As EpsDotNet.dispenser
   Set dispenser = dispensers(dispensersCount)
   
   Dim pumps As Variant
   pumps = dispenser.GetPumps
   
   Dim pumpsCount As Integer
   For pumpsCount = 0 To UBound(pumps)
     Dim pump As EpsDotNet.pump
     Set pump = pumps(pumpsCount)
   Next pumpsCount
   MsgBox "Site " + CStr(siteID) + " for Dispenser " + CStr(dispenser.dispenserId) + " Pumps Count " + CStr(pumpsCount)
   Set dispenser = Nothing
 Next dispensersCount
 MsgBox "Site " + CStr(siteID) + " Dispenser Count " + CStr(dispensersCount)
End Function

Function RemovePumpFromExistingDispenser(controller As EpsDotNet.FuelSystemController)

   Dim allDispensers As Variant
   allDispensers = controller.GetDispensers(33)
    
   Dim dispensersCount As Integer
   For dispensersCount = 0 To UBound(allDispensers)
    Dim dispenser As EpsDotNet.dispenser
    Set dispenser = allDispensers(dispensersCount)
    If dispenser.DispenserNumber = 14 Then
      Dim pumps As Variant
      pumps = dispenser.GetPumps
     
      Dim pumpsCount As Integer
      For pumpsCount = 0 To UBound(pumps)
     
          Dim pump As EpsDotNet.pump
          Set pump = pumps(pumpsCount)
          If (pump.ForecourtPumpNumber = 3) Then
              dispenser.RemovePump pump
          End If
      Next pumpsCount
      Dim returnDispenser As EpsDotNet.dispenser
      Set returnDispenser = controller.CommitChangesDispenser(dispenser)
       
   End If
   Set dispenser = Nothing
  Next dispensersCount

End Function

Function RemoveDispenser(controller As EpsDotNet.FuelSystemController)

'Dim siteID As Long
'siteID = 33

Dim dispensers As Variant

dispensers = controller.GetDispensers(siteID)

Dim dispensersCount As Integer
    For dispensersCount = 0 To UBound(dispensers)
       Dim dispenser As EpsDotNet.dispenser
       Set dispenser = dispensers(dispensersCount)
       If (dispenser.DispenserNumber = 17) Then
            controller.RemoveDispenser dispenser.dispenserId
       End If
    Next dispensersCount
   
End Function

'fuelSystem
' Need to know valid  SiteID, Grade ID, and use GetHoseportion function to get to Hoseportion to attach
Function AddNewFuelSystemWithGradeTankHosePortion(controller As EpsDotNet.FuelSystemController)

    'Dim siteID As Long
    'siteID = 2
    
    Dim tankExist As EpsDotNet.tank
    Set tankExist = New EpsDotNet.tank
    tankExist.TankId = 18
    
    Dim gradeExist As EpsDotNet.grade
    Set gradeExist = New EpsDotNet.grade
    gradeExist.GradeId = 4

    Dim hosePortionExist As EpsDotNet.hosePortion
    Set hosePortionExist = GetHoseportion(controller)
    
    Dim hoseSet As EpsDotNet.hoseSet
    Set hoseSet = New EpsDotNet.hoseSet
    hoseSet.AttachHosePortion hosePortionExist
    
    Dim fuelSystem As EpsDotNet.fuelSystem
    Set fuelSystem = New EpsDotNet.fuelSystem

    fuelSystem.siteID = siteID
    fuelSystem.ufi = "ufi"
    fuelSystem.SetFromdate (Now())
    fuelSystem.AttachGrade gradeExist
    fuelSystem.AttachTank tankExist
    fuelSystem.AddHoseSet hoseSet
    fuelSystem.Active = True

    Dim returnFuelSystem As EpsDotNet.fuelSystem
    Set returnFuelSystem = controller.CommitChangesFuelSystem(fuelSystem)
   
End Function

'To get hose portion to attach to fuelsystem
Function GetHoseportion(controller As EpsDotNet.FuelSystemController) As EpsDotNet.hosePortion
    'Dim siteID As Long
    'siteID = 33
    
    Dim dispenserNo As Long
    dispenserNo = 15
    Dim ForecourtPumpNumber As Long
    ForecourtPumpNumber = 1
    Dim ForecourtHoseNumber As Long
    ForecourtHoseNumber = 1
    
    Dim dispensers As Variant
    dispensers = controller.GetDispensers(siteID)

    Dim dispensersCount As Integer
    For dispensersCount = 0 To UBound(dispensers)
      Dim dispenser As EpsDotNet.dispenser
      Set dispenser = dispensers(dispensersCount)
      If dispenser.DispenserNumber = dispenserNo Then
        Dim pumps As Variant
        pumps = dispenser.GetPumps
        Dim pumpsCount As Integer
        For pumpsCount = 0 To UBound(pumps)
          Dim pump As EpsDotNet.pump
          Set pump = pumps(pumpsCount)
          If pump.ForecourtPumpNumber = ForecourtPumpNumber Then
            Dim hoses As Variant
            hoses = pump.GetHoses
            Dim hosesCount As Integer
            For hosesCount = 0 To UBound(hoses)
             Dim hose As EpsDotNet.hose
             Set hose = hoses(hosesCount)
             If hose.ForecourtHoseNumber = ForecourtHoseNumber Then
                Dim hosePortions As Variant
                hosePortions = hose.GetHosePortions
                Set GetHoseportion = hosePortions(0)
             End If
             
            Next hosesCount
          End If
        Next pumpsCount
      End If
      
      Set dispenser = Nothing
    Next dispensersCount
End Function

Function GetActiveFuelSystem(controller As EpsDotNet.FuelSystemController)

    'Dim siteID As Long
    'siteID = 2
    
    Dim ufi As String
    ufi = "test22"
    
    Dim fuelSystem As EpsDotNet.fuelSystem
    Set fuelSystem = controller.GetActiveFuelSystem(siteID, ufi)
    If Not fuelSystem Is Nothing Then
        MsgBox "Site= " + CStr(siteID) + ",UFi= " + ufi + ", FuelSytemId= " + CStr(fuelSystem.FuelSystemId)
    End If
    
End Function

Function UpdateFuelSystemByAddingTankGrade(controller As EpsDotNet.FuelSystemController)
   
   Dim ufi As String
   ufi = "test22"
   
   'Dim siteID As Long
   'siteID = 33
   
   Dim companyID As Long
   companyID = 3
   
   Dim fuelSystemToUpdate As EpsDotNet.fuelSystem
   Set fuelSystemToUpdate = controller.GetActiveFuelSystem(siteID, ufi)
   
   'Modify FuelSystem
   fuelSystemToUpdate.SetFromdate (Now())
   
   'Get existing tanks and attach first tank to fuelsystem
   Dim tankExist As EpsDotNet.tank
   Dim tankVariant As Variant
    tankVariant = controller.GetTanks(siteID)
   If (variantCount(tankVariant) > 4) Then
      Set tankExist = tankVariant(4)
      fuelSystemToUpdate.AttachTank tankExist
   End If
   
   'Get existing grades and attach first grade to fuelsystem
   Dim gradeExist As EpsDotNet.grade
   Dim gradeVariant As Variant
    gradeVariant = controller.GetGrades(companyID)
   If (variantCount(gradeVariant) > 4) Then
      Set gradeExist = gradeVariant(4)
      fuelSystemToUpdate.AttachGrade gradeExist
   End If
   
   controller.CommitChangesFuelSystem fuelSystemToUpdate

End Function

Function CloneFuelSystem(controller As EpsDotNet.FuelSystemController)
   
   Dim ufi As String
   ufi = "ufi"
   
   'Dim siteID As Long
   'siteID = 33
   
   Dim companyID As Long
   companyID = 3
   
   'Existing tank already attached to clone Fuel system
   Dim tankAttachedToClone As EpsDotNet.tank
   Set tankAttachedToClone = New EpsDotNet.tank
   tankAttachedToClone.ForecourtTankNo = 5
   
   'Existing grade already attached to clone Fuel system
   Dim gradeAttachedToClone As EpsDotNet.grade
   Set gradeAttachedToClone = New EpsDotNet.grade
   gradeAttachedToClone.gradeName = "Diesel"
   
   Dim ClnFuelSystem As EpsDotNet.fuelSystem
   Set ClnFuelSystem = controller.CloneFuelSystem(siteID, ufi)
   
   'Modify Clone FuelSystem
   ClnFuelSystem.SetFromdate (Now())
   ClnFuelSystem.DetachTank tankAttachedToClone
   ClnFuelSystem.DetachGrade gradeAttachedToClone
   
   Dim hoseSetArray As Variant
   hoseSetArray = ClnFuelSystem.GetHoseSet
   'Remove all hoseset
   Dim hoseSetArrayCount As Integer
   For hoseSetArrayCount = 0 To UBound(hoseSetArray)
     ClnFuelSystem.RemoveHoseSet hoseSetArray(hoseSetArrayCount)
   Next hoseSetArrayCount
   
   'Get existing tanks and attach first tank to fuelsystem
   Dim tankExist As EpsDotNet.tank
   Dim tankVariant As Variant
    tankVariant = controller.GetTanks(siteID)
   If (variantCount(tankVariant) > 3) Then
      Set tankExist = tankVariant(3)
      ClnFuelSystem.AttachTank tankExist
   End If
   
   'Get existing grades and attach first grade to fuelsystem
   Dim gradeExist As EpsDotNet.grade
   Dim gradeVariant As Variant
    gradeVariant = controller.GetGrades(companyID)
   If (variantCount(gradeVariant) > 3) Then
      Set gradeExist = gradeVariant(3)
      ClnFuelSystem.AttachGrade gradeExist
   End If
   
   controller.CommitChangesFuelSystem ClnFuelSystem

End Function

'Exception Handling

Function AddDuplicatePumps(variantObj As Variant)
    Dim pump1 As EpsDotNet.pump
    Set pump1 = New EpsDotNet.pump
    pump1.ForecourtPumpNumber = 7
    
    Dim pump2 As EpsDotNet.pump
    Set pump2 = New EpsDotNet.pump
    pump2.ForecourtPumpNumber = 8
  
    Dim dispenser As EpsDotNet.dispenser
    Set dispenser = New EpsDotNet.dispenser
    dispenser.DispenserName = "Dispenser Name"
    dispenser.siteID = 33
    dispenser.SerialNumber = "SER 565"
    dispenser.DispenserNumber = 15
    
    Call AddPump(dispenser, pump1)
    Call AddPump(dispenser, pump1)
    Call AddPump(dispenser, pump2)
    
    Dim pumpsVariant As Variant
    pumpsVariant = dispenser.GetPumps
    MsgBox CStr(variantCount(pumpsVariant))
    
End Function

Public Sub AddPump(dispenser As EpsDotNet.dispenser, pump As EpsDotNet.pump)
   On Error GoTo ErrorHandler
    dispenser.AddPump pump
Exit Sub
ErrorHandler:
   MsgBox Err.Description
   Resume Next
End Sub

