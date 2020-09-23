VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clip"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   531
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelRect As LINE
Dim Face(2) As LINE
Dim Poly()  As LINE

Private Sub Form_Load()

    Face(0).P1.X = 280:    Face(0).P1.Y = 120
    Face(0).P2.X = 190:    Face(0).P2.Y = 240
    Face(1).P1.X = 190:    Face(1).P1.Y = 240
    Face(1).P2.X = 350:    Face(1).P2.Y = 240
    Face(2).P1.X = 350:    Face(2).P1.Y = 240
    Face(2).P2.X = 280:    Face(2).P2.Y = 120

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.DrawMode = 6
    SelRect.P1.X = X
    SelRect.P1.Y = Y
    SelRect.P2 = SelRect.P1
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button Then
        DrawLine SelRect.P1, SelRect.P2, vbBlack, True
        SelRect.P2.X = X
        SelRect.P2.Y = Y
        DrawLine SelRect.P1, SelRect.P2, vbBlack, True
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    Dim idx As Integer
    
    Me.Cls
    Me.DrawMode = 13
    Me.DrawWidth = 1
'Draw Rectangle
    SelRect.P2.X = X
    SelRect.P2.Y = Y
    If (SelRect.P1.X > SelRect.P2.X) Then Call SwapLong(SelRect.P1.X, SelRect.P2.X)
    If (SelRect.P1.Y > SelRect.P2.Y) Then Call SwapLong(SelRect.P1.Y, SelRect.P2.Y)
    Rect.T.P1 = SelRect.P1: Rect.T.P2.X = SelRect.P2.X: Rect.T.P2.Y = SelRect.P1.Y
    Rect.R.P1 = Rect.T.P2:  Rect.R.P2 = SelRect.P2
    Rect.D.P1 = Rect.R.P2:  Rect.D.P2.X = SelRect.P1.X: Rect.D.P2.Y = SelRect.P2.Y
    Rect.L.P1 = Rect.D.P2:  Rect.L.P2 = Rect.T.P1
    DrawLine Rect.T.P1, Rect.T.P2, vbGreen
    DrawLine Rect.R.P1, Rect.R.P2, vbGreen
    DrawLine Rect.D.P1, Rect.D.P2, vbGreen
    DrawLine Rect.L.P1, Rect.L.P2, vbGreen
'Draw Triangle
    For idx = 0 To 2
        DrawLine Face(idx).P1, Face(idx).P2, vbMagenta
    Next idx
'Draw Polygon
    ReDim Poly(0)
    Me.DrawWidth = 2
    Call ClipTriangle(Face(0).P1, Face(1).P1, Face(2).P1, Poly)
    If UBound(Poly) Then
        For idx = 0 To UBound(Poly)
            DrawLine Poly(idx).P1, Poly(idx).P2, vbCyan
        Next idx
    End If
    
'    If UBound(Pts) Then
'        For idx = 1 To UBound(Pts)
'            Me.PSet (Pts(idx).X, Pts(idx).Y), vbBlue
'        Next
'        Me.PSet (CenterPoint.X, CenterPoint.Y), vbRed
'    End If
    
End Sub

Private Sub DrawLine(P1 As POINTAPI, P2 As POINTAPI, Color As ColorConstants, Optional Box As Boolean = False)
    
    If Box = False Then
        Me.Line (P1.X, P1.Y)-(P2.X, P2.Y), Color
    Else
        Me.Line (P1.X, P1.Y)-(P2.X, P2.Y), Color, BF
    End If
    
End Sub

Private Sub SwapLong(V1 As Long, V2 As Long)
    
    Dim V3 As Long
    
    V3 = V1: V1 = V2: V2 = V3

End Sub
