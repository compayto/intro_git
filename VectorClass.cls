VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VectorClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private valX As Double
Private valY As Double
Private valZ As Double

Property Get x() As Double
    x = valX
End Property

Property Let x(inputX As Double)
    valX = inputX
End Property

Property Get y() As Double
    y = valY
End Property

Property Let y(inputY As Double)
    valY = inputY
End Property

Property Get z() As Double
    z = valZ
End Property

Property Let z(inputZ As Double)
    valZ = inputZ
End Property
