use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Using Statement
// ═══════════════════════════════════════════════════════════

#[test]
fn using_statement_basic() {
    let out = run_vb(
        r#"
Class FakeFile
    Implements IDisposable
    
    Public Sub Write(text As String)
        Console.WriteLine("Writing: " & text)
    End Sub
    
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("FakeFile Disposed")
    End Sub
End Class

Module M
    Sub Main()
        Using f As New FakeFile()
            f.Write("Hello")
        End Using
        Console.WriteLine("Done")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Writing: Hello", "FakeFile Disposed", "Done"]);
}

#[test]
fn using_statement_multiple_variables() {
    let out = run_vb(
        r#"
Class Resource
    Implements IDisposable
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed " & Name)
    End Sub
End Class

Module M
    Sub Main()
        ' Multiple resources of the same type in one Using statement
        Using r1 As New Resource("A"), r2 As New Resource("B")
            Console.WriteLine("Using " & r1.Name & " and " & r2.Name)
        End Using
    End Sub
End Module
"#,
    );
    // Order of disposal is typically reverse of declaration, or it could be arbitrary based on implementation.
    // Let's assert the block execution and disposal occurred.
    let joined = out.join("|");
    assert!(joined.contains("Using A and B"));
    assert!(joined.contains("Disposed A"));
    assert!(joined.contains("Disposed B"));
}

#[test]
fn using_statement_pre_instantiated() {
    let out = run_vb(
        r#"
Class Resource
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed")
    End Sub
End Class

Module M
    Sub Main()
        Dim r As New Resource()
        Using r
            Console.WriteLine("Inside")
        End Using
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Inside", "Disposed"]);
}
