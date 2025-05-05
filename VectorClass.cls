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

Property Get norm() As Double
    norm = Sqr(x * x + y * y + z * z)
End Property

Public Function DotProduct(toVector As VectorClass) As Double
    DotProduct = x * toVector.x + y * toVector.y + z * toVector.z
End Function

Public Function CrossProduct(toVector As VectorClass) As VectorClass
    Set CrossProduct = New VectorClass
    CrossProduct.x = y * toVector.z - z * toVector.y
    CrossProduct.y = z * toVector.x - x * toVector.z
    CrossProduct.z = z * toVector.y - y * toVector.x
End Function

Public Function Normalization() As VectorClass
    Set Normalization = New VectorClass
    If norm <> 0 Then
        Normalization.x = x / norm
        Normalization.y = y / norm
        Normalization.z = z / norm
    End If
End Function

Property Let setVectorFromStr(str As String)
    Dim buf As Variant: buf = Split(Replace(Replace(str, "(", ""), ")", ""), ",")
    If UBound(buf) - LBound(buf) + 1 = 3 Then
        valX = buf(0)
        valY = buf(1)
        valZ = buf(2)
    Else
        Call errorComand(-1)
    End If
End Property

Public Sub debug_()
    
    Debug.Print x
    Debug.Print y
    Debug.Print z

End Sub

Private Sub errorComand(errorcode As Long)
    Dim errorMsg As String
    Select Case errorcode
    Case -1: errorMsg = "3éüå≥Ç∂Ç·Ç»Ç¢Ç≈Ç∑"
    Case Else: errorMsg = "ïsñæÇ»ÉGÉâÅ[Ç≈Ç∑"
    End Select
    Call MsgBox(errorMsg, vbInformation)
End Sub

