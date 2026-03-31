Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class PaintForm
    Private Enum DrawMode
        Doodle
        Circle
        Rectangle
    End Enum

    Private _currentMode As DrawMode = DrawMode.Doodle
    Private _isDrawing As Boolean = False
    Private _startPoint As Point
    Private _lastPoint As Point
    Private _shapes As New List(Of Shape)

    Private Sub btnDoodle_Click(sender As Object, e As EventArgs) Handles btnDoodle.Click
        _currentMode = DrawMode.Doodle
    End Sub

    Private Sub btnCircle_Click(sender As Object, e As EventArgs) Handles btnCircle.Click
        _currentMode = DrawMode.Circle
    End Sub

    Private Sub btnRectangle_Click(sender As Object, e As EventArgs) Handles btnRectangle.Click
        _currentMode = DrawMode.Rectangle
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        _shapes.Clear()
        pnlCanvas.Invalidate()
    End Sub

    Private Sub pnlCanvas_MouseDown(sender As Object, e As MouseEventArgs) Handles pnlCanvas.MouseDown
        _isDrawing = True
        _startPoint = e.Location
        _lastPoint = e.Location
        
        If _currentMode = DrawMode.Doodle Then
            Dim doodle As New DoodleShape()
            doodle.Points.Add(e.Location)
            _shapes.Add(doodle)
        End If
    End Sub

    Private Sub pnlCanvas_MouseMove(sender As Object, e As MouseEventArgs) Handles pnlCanvas.MouseMove
        If _isDrawing Then
            If _currentMode = DrawMode.Doodle Then
                Dim doodle = DirectCast(_shapes(_shapes.Count - 1), DoodleShape)
                doodle.Points.Add(e.Location)
            End If
            _lastPoint = e.Location
            pnlCanvas.Invalidate()
        End If
    End Sub

    Private Sub pnlCanvas_MouseUp(sender As Object, e As MouseEventArgs) Handles pnlCanvas.MouseUp
        If _isDrawing Then
            _isDrawing = False
            _lastPoint = e.Location
            
            If _currentMode = DrawMode.Circle Then
                _shapes.Add(New CircleShape(_startPoint, _lastPoint))
            ElseIf _currentMode = DrawMode.Rectangle Then
                _shapes.Add(New RectangleShape(_startPoint, _lastPoint))
            End If
            
            pnlCanvas.Invalidate()
        End If
    End Sub

    Private Sub pnlCanvas_Paint(sender As Object, e As PaintEventArgs) Handles pnlCanvas.Paint
        Dim g = e.Graphics
        For Each shape In _shapes
            shape.Draw(g)
        Next
        
        ' Draw current preview
        If _isDrawing Then
            If _currentMode = DrawMode.Circle Then
                Dim rect = GetRect(_startPoint, _lastPoint)
                g.DrawEllipse(Pens.Black, rect)
            ElseIf _currentMode = DrawMode.Rectangle Then
                Dim rect = GetRect(_startPoint, _lastPoint)
                g.DrawRectangle(Pens.Black, rect)
            End If
        End If
    End Sub

    Private Function GetRect(p1 As Point, p2 As Point) As Rectangle
        Return New Rectangle(Math.Min(p1.X, p2.X), Math.Min(p1.Y, p2.Y), Math.Abs(p1.X - p2.X), Math.Abs(p1.Y - p2.Y))
    End Function

    ' Helper classes for shapes
    Public MustInherit Class Shape
        Public MustOverride Sub Draw(g As Graphics)
    End Class

    Public Class DoodleShape
        Inherits Shape
        Public Property Points As New List(Of Point)
        Public Overrides Sub Draw(g As Graphics)
            If Points.Count > 1 Then
                For i As Integer = 0 To Points.Count - 2
                    g.DrawLine(Pens.Black, Points(i), Points(i + 1))
                Next
            End If
        End Sub
    End Class

    Public Class CircleShape
        Inherits Shape
        Public Property StartPoint As Point
        Public Property EndPoint As Point
        Public Sub New(s As Point, e As Point)
            StartPoint = s
            EndPoint = e
        End Sub
        Public Overrides Sub Draw(g As Graphics)
            Dim rect = New Rectangle(Math.Min(StartPoint.X, EndPoint.X), Math.Min(StartPoint.Y, EndPoint.Y), Math.Abs(StartPoint.X - EndPoint.X), Math.Abs(StartPoint.Y - EndPoint.Y))
            g.DrawEllipse(Pens.Black, rect)
        End Sub
    End Class

    Public Class RectangleShape
        Inherits Shape
        Public Property StartPoint As Point
        Public Property EndPoint As Point
        Public Sub New(s As Point, e As Point)
            StartPoint = s
            EndPoint = e
        End Sub
        Public Overrides Sub Draw(g As Graphics)
            Dim rect = New Rectangle(Math.Min(StartPoint.X, EndPoint.X), Math.Min(StartPoint.Y, EndPoint.Y), Math.Abs(StartPoint.X - EndPoint.X), Math.Abs(StartPoint.Y - EndPoint.Y))
            g.DrawRectangle(Pens.Black, rect)
        End Sub
    End Class
End Class
