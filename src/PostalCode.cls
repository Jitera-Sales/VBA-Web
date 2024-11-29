VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostalCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

' Attributes for the PostalCode entity
Private postalCode As String
Private city As String
Private region As String

' Property to get and set PostalCode
Public Property Get PostalCode() As String
    PostalCode = postalCode
End Property

Public Property Let PostalCode(value As String)
    postalCode = value
End Property

' Property to get and set City
Public Property Get City() As String
    City = city
End Property

Public Property Let City(value As String)
    city = value
End Property

' Property to get and set Region
Public Property Get Region() As String
    Region = region
End Property

Public Property Let Region(value As String)
    region = value
End Property

' Method to initialize the PostalCode entity
Public Sub Initialize(ByVal pCode As String, ByVal pCity As String, ByVal pRegion As String)
    postalCode = pCode
    city = pCity
    region = pRegion
End Sub

' Method to display the postal code information
Public Sub DisplayInfo()
    Debug.Print "Postal Code: " & postalCode
    Debug.Print "City: " & city
    Debug.Print "Region: " & region
End Sub