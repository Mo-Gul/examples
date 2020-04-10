VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AbstractCarFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("Examples.ReadOnlyEverywhere.AbstractFactory")
'@PredeclaredId
'@Exposed
Option Explicit
Implements ICarFactory

'NOTE: This class doesn't seem to be used anywhere.
'      I also couldn't find this in the article.
'      Do I miss something or can this class be deleted?
'      Please note that in the code the commented lines need to be commented
'      to avoid comiler errors.
Private Function ICarFactory_Create(ByVal carMake As Long, ByVal carModel As String, ByVal carManufacturer As String) As ICar
    Dim result As ICar
    Set result = New ReadOnlyCar '<~ we could use a strategy pattern here and output a different implementation depending on parameters
'    result.Make = carMake
'    result.Model = carModel
'    result.Manufacturer = carManufacturer
    Set ICarFactory_Create = result
End Function
