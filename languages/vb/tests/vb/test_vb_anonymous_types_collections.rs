use super::helpers::run_vb;

#[test]
fn anonymous_types_collections() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Array of anonymous types
        Dim arr = {
            New With { Key .Id = 1, .Name = "Alice" },
            New With { Key .Id = 2, .Name = "Bob" }
        }
        
        Console.WriteLine(arr(0).Name)
        Console.WriteLine(arr.Length)
        
        ' Collection initializer with anonymous types
        Dim list As New System.Collections.Generic.List(Of Object) From {
            New With { .Value = 10 },
            New With { .Value = 20 }
        }
        Console.WriteLine(list.Count)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "2", "2"]);
}

#[test]
fn anonymous_types_equality() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Two anonymous types with the same Key properties are considered equal
        Dim a1 = New With { Key .X = 1, Key .Y = 2, .Z = 3 }
        Dim a2 = New With { Key .X = 1, Key .Y = 2, .Z = 4 }
        Dim a3 = New With { Key .X = 2, Key .Y = 2, .Z = 3 }
        
        Console.WriteLine(a1.Equals(a2))
        Console.WriteLine(a1.Equals(a3))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}
