VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISimplerCarFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Examples.ReadOnlyEverywhere.AbstractFactory")
'@Interface
Option Explicit

Public Function Create(ByVal carModel As String) As ICar
End Function

