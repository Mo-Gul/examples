Attribute VB_Name = "CarFactoryExample"
'@Folder("VBAProject")
Option Explicit

'NOTE: Either this module was missing in this project or
'      it was intended to be in the 'ReferencingProject.xlsm'.
'      In the later case please mention that in the article
'      (and continue to read comments there)
Public Sub DoSomething()
    Dim myCar As Car
    Set myCar = CarFactory.Create(2016, "Civic", "Honda")
    
    MsgBox "We have a " & myCar.Make & " " & myCar.Manufacturer & " " & myCar.Model & " here."
    'these assignments are illegal here, code won't compile if they're uncommented:
    'myCar.Make = 2014
    'myCar.Model = "Fit"
    
End Sub
