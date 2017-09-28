VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAdoDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private sConnectString As String
Private oConnect As ADODB.Connection

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    On Error Resume Next

    If Not (oConnect Is Nothing) Then
        If (oConnect.State And 1) = 1 Then oConnect.Close
        Set oConnect = Nothing
    End If
End Sub

Public Property Let ConnectionString(ByVal stConn As String)
    sConnectString = stConn
End Property
Private Property Get ConnectionString() As String
    ConnectionString = sConnectString
End Property
