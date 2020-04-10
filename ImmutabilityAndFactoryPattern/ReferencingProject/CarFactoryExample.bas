Attribute VB_Name = "CarFactoryExample"
'@Folder("VBAProject")
Option Explicit

'if this should work from this file (see comment in same module in
''CarsProject.xlsm') then I guess a reference to the other project is missing,
'right?
Public Sub DoSomething()
    Dim myCar As Car '<~ DoSomething is coupled with the Car class here
    Set myCar = CarFactory.Create(2016, "Civic", "Honda") '<~ DoSomething is also coupled with the CarFactory class
    
    MsgBox "We have a " & myCar.Make & " " & myCar.Manufacturer & " " & myCar.Model & " here."
    'these assignments are illegal here, code won't compile if they're uncommented:
    'myCar.Make = 2014
    'myCar.Model = "Fit"
    
End Sub

