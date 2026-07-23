use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: EventHandler(Of TEventArgs) & Custom EventArgs
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_event_handler_generic_custom_event_args() {
    let src = r#"
Imports System

Class OrderEventArgs
    Inherits EventArgs
    Public ReadOnly OrderId As Integer
    Public Sub New(id As Integer)
        Me.OrderId = id
    End Sub
End Class

Class OrderProcessor
    Public Event OrderProcessed As EventHandler(Of OrderEventArgs)

    Public Sub Process(id As Integer)
        RaiseEvent OrderProcessed(Me, New OrderEventArgs(id))
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New OrderProcessor()
        AddHandler p.OrderProcessed, Sub(sender, e)
            Console.WriteLine("Order: " & e.OrderId)
        End Sub
        p.Process(1001)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Order: 1001"]);
}

#[test]
fn test_vb_event_handler_empty_event_args() {
    let src = r#"
Imports System

Class Button
    Public Event Click As EventHandler

    Public Sub PerformClick()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim btn As New Button()
        AddHandler btn.Click, Sub(sender, e)
            Console.WriteLine("Clicked")
        End Sub
        btn.PerformClick()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Clicked"]);
}

#[test]
fn test_vb_event_handler_cancel_event_args() {
    let src = r#"
Imports System
Imports System.ComponentModel

Class Document
    Public Event Closing As EventHandler(Of CancelEventArgs)

    Public Function TryClose() As Boolean
        Dim args As New CancelEventArgs()
        RaiseEvent Closing(Me, args)
        Return Not args.Cancel
    End Function
End Class

Module Program
    Sub Main()
        Dim doc As New Document()
        AddHandler doc.Closing, Sub(sender, e)
            e.Cancel = True
        End Sub
        Console.WriteLine(doc.TryClose())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}
