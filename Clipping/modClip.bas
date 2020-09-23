Attribute VB_Name = "modClip"
Option Explicit

Const Pi            As Single = 3.141592
Const HalfPi        As Single = 1.570796
Const ApproachVal   As Single = 0.000001

Public Type POINTAPI
    X       As Long
    Y       As Long
End Type

Public Type LINE
    P1      As POINTAPI
    P2      As POINTAPI
End Type

Public Type CANVASRECT
    T       As LINE 'T op
    R       As LINE 'R ight
    D       As LINE 'D own
    L       As LINE 'L eft
End Type

Private Type ORDER
    Value   As Single
    Index   As Integer
End Type

Public Rect         As CANVASRECT
Public Pts()        As POINTAPI
Public CenterPoint  As POINTAPI
Private ordList()   As ORDER

Public Sub ClipTriangle(P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI, Poly() As LINE)

    ReDim Pts(0)
    IsInTriangle P1, P2, P3, Rect.T.P1.X, Rect.T.P1.Y
    IsInTriangle P1, P2, P3, Rect.R.P1.X, Rect.R.P1.Y
    IsInTriangle P1, P2, P3, Rect.D.P1.X, Rect.D.P1.Y
    IsInTriangle P1, P2, P3, Rect.L.P1.X, Rect.L.P1.Y
    IsInRectangle P1
    IsInRectangle P2
    IsInRectangle P3
    RectLineIntersection P1, P2
    RectLineIntersection P2, P3
    RectLineIntersection P3, P1
    If UBound(Pts) Then
        GetCenterPoint
        GetAngles
        SortAngles 0, UBound(ordList)
        MakePolygon Poly
    End If
    
End Sub

Private Function IsInTriangle(P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI, PX As Long, PY As Long) As Boolean

    Dim Val1 As Single
    Dim Val2 As Single
    Dim Val3 As Single
       
    Val1 = (P1.X - PX) * (P2.Y - PY) - (P2.X - PX) * (P1.Y - PY)
    Val2 = (P2.X - PX) * (P3.Y - PY) - (P3.X - PX) * (P2.Y - PY)
    Val3 = (P3.X - PX) * (P1.Y - PY) - (P1.X - PX) * (P3.Y - PY)
    
    If (Val1 > 0 And Val2 > 0 And Val3 > 0) Or _
       (Val1 < 0 And Val2 < 0 And Val3 < 0) Then
        AddPoint PX, PY
        IsInTriangle = True
    End If
    
End Function

Private Function IsInRectangle(P1 As POINTAPI) As Boolean

    If P1.X > Rect.L.P1.X And P1.X < Rect.R.P1.X And _
       P1.Y > Rect.T.P1.Y And P1.Y < Rect.D.P1.Y Then
       AddPoint P1.X, P1.Y
       IsInRectangle = True
    End If

End Function

Private Sub AddPoint(X As Long, Y As Long)
        
    ReDim Preserve Pts(UBound(Pts) + 1)
    Pts(UBound(Pts)).X = X
    Pts(UBound(Pts)).Y = Y
        
End Sub

Private Function LineLineIntersection(P1X1 As Long, P1Y1 As Long, _
                                      P1X2 As Long, P1Y2 As Long, _
                                      P2X1 As Long, P2Y1 As Long, _
                                      P2X2 As Long, P2Y2 As Long) As Boolean
                                      
    Dim D1X As Long
    Dim D1Y As Long
    Dim D2X As Long
    Dim D2Y As Long
    Dim T   As Single
    Dim S   As Single

    D1X = P1X2 - P1X1
    D1Y = P1Y2 - P1Y1
    D2X = P2X2 - P2X1
    D2Y = P2Y2 - P2Y1

    If (D2X * D1Y - D2Y * D1X) = 0 Then
        ' The lines are parallel.
        LineLineIntersection = False
        Exit Function
    End If

    S = (D1X * (P2Y1 - P1Y1) + D1Y * (P1X1 - P2X1)) / (D2X * D1Y - D2Y * D1X)
    T = (D2X * (P1Y1 - P2Y1) + D2Y * (P2X1 - P1X1)) / (D2Y * D1X - D2X * D1Y)
    LineLineIntersection = (S >= 0# And S <= 1# And T >= 0# And T <= 1#)

    ' If it exists, the point of intersection is:
    If LineLineIntersection Then AddPoint P1X1 + T * D1X, P1Y1 + T * D1Y
     
End Function

Private Sub RectLineIntersection(P1 As POINTAPI, P2 As POINTAPI)

    LineLineIntersection Rect.T.P1.X, Rect.T.P1.Y, Rect.T.P2.X, Rect.T.P2.Y, P1.X, P1.Y, P2.X, P2.Y
    LineLineIntersection Rect.L.P1.X, Rect.L.P1.Y, Rect.L.P2.X, Rect.L.P2.Y, P1.X, P1.Y, P2.X, P2.Y
    LineLineIntersection Rect.R.P1.X, Rect.R.P1.Y, Rect.R.P2.X, Rect.R.P2.Y, P1.X, P1.Y, P2.X, P2.Y
    LineLineIntersection Rect.D.P1.X, Rect.D.P1.Y, Rect.D.P2.X, Rect.D.P2.Y, P1.X, P1.Y, P2.X, P2.Y

End Sub

Private Sub GetCenterPoint()
    
    Dim idx As Integer
    Dim X As Long
    Dim Y As Long
        
    For idx = 1 To UBound(Pts)
        X = X + Pts(idx).X
        Y = Y + Pts(idx).Y
    Next
    CenterPoint.X = CLng(X / UBound(Pts))
    CenterPoint.Y = CLng(Y / UBound(Pts))

End Sub

Private Function Angle(P1 As POINTAPI, P2 As POINTAPI) As Single
'Radian
    Angle = (Atn(Div((P2.X - P1.X), (P2.Y - P1.Y)))) + HalfPi
    If P2.Y >= P1.Y Then Angle = Pi + Angle

End Function

Private Function Div(R1 As Single, ByVal R2 As Single) As Single
    
    If R2 = 0 Then R2 = ApproachVal
    Div = R1 / R2

End Function

Private Sub GetAngles()
    
    Dim idx As Integer

    ReDim ordList(UBound(Pts) - 1)
    
    For idx = 1 To UBound(Pts)
        ordList(idx - 1).Value = Angle(CenterPoint, Pts(idx))
        ordList(idx - 1).Index = idx
    Next

End Sub

Private Sub SortAngles(ByVal First As Long, ByVal Last As Long)

    Dim FirstIdx    As Long
    Dim MidIdx      As Long
    Dim LastIdx     As Long
    Dim MidVal      As Single
    Dim ordTemp     As ORDER
    
    If (First < Last) Then
            MidIdx = (First + Last) * 0.5
            MidVal = ordList(MidIdx).Value
            FirstIdx = First
            LastIdx = Last
            Do
                Do While ordList(FirstIdx).Value < MidVal
                    FirstIdx = FirstIdx + 1
                Loop
                Do While ordList(LastIdx).Value > MidVal
                    LastIdx = LastIdx - 1
                Loop
                If (FirstIdx <= LastIdx) Then
                    ordTemp = ordList(LastIdx)
                    ordList(LastIdx) = ordList(FirstIdx)
                    ordList(FirstIdx) = ordTemp
                    FirstIdx = FirstIdx + 1
                    LastIdx = LastIdx - 1
                End If
            Loop Until FirstIdx > LastIdx

            If (LastIdx <= MidIdx) Then
                SortAngles First, LastIdx
                SortAngles FirstIdx, Last
            Else
                SortAngles FirstIdx, Last
                SortAngles First, LastIdx
            End If
    End If

End Sub

Private Sub MakePolygon(Poly() As LINE)
    
    Dim idx As Integer
    
    ReDim Poly(UBound(ordList))
    
    For idx = 0 To UBound(ordList) - 1
        Poly(idx).P1 = Pts(ordList(idx).Index)
        Poly(idx).P2 = Pts(ordList(idx + 1).Index)
    Next
    Poly(UBound(ordList)).P1 = Pts(ordList(UBound(ordList)).Index)
    Poly(UBound(ordList)).P2 = Pts(ordList(0).Index)

End Sub
