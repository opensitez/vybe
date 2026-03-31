Imports System.Windows.Forms
Imports System.Drawing

Public Class TestForm
    Inherits Form

    Public WithEvents btn As Button

    Public Sub New()
        btn = New Button()
        btn.Name = "btn1"
        Console.WriteLine("Button created and assigned.")
    End Sub

    Private Sub HandleClick() Handles btn.Click
        Console.WriteLine("Button Clicked (Handler Invoked)!")
    End Sub

    Public Sub TestDrawing()
        Console.WriteLine("Testing Drawing...")
        Dim g As Graphics = Me.CreateGraphics()
        Console.WriteLine("Graphics created.")
        Dim p As New Pen(Color.Red, 5)
        g.DrawLine(p, 10, 10, 100, 100)
        Console.WriteLine("DrawLine executed.")
        
        g.DrawRectangle(p, 50, 50, 20, 20)
        Console.WriteLine("DrawRectangle executed.")
        
        Dim b As New SolidBrush(Color.Blue)
        g.FillRectangle(b, 0, 0, 20, 20)
        Console.WriteLine("FillRectangle executed.")
    End Sub
End Class

Module Program
    Sub Main()
        Console.WriteLine("Starting Test...")
        Dim f As New TestForm()
        f.TestDrawing()
        
        ' Note: We cannot trigger the actual Click event here easily as it requires UI event loop.
        ' But if we reached here, wiring code ran without crashing.
        Console.WriteLine("Test Finished.")
    End Sub
End Module
