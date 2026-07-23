use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Anonymous Types Key Properties & Equality
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_anonymous_type_key_property_equals() {
    let src = r#"
Module Program
    Sub Main()
        Dim p1 = New With {Key .Id = 1, .Name = "Alice"}
        Dim p2 = New With {Key .Id = 1, .Name = "Bob"} ' Same Key, different non-key
        Dim p3 = New With {Key .Id = 2, .Name = "Alice"}

        Console.WriteLine(p1.Equals(p2))
        Console.WriteLine(p1.Equals(p3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_anonymous_type_get_hash_code() {
    let src = r#"
Module Program
    Sub Main()
        Dim p1 = New With {Key .Id = 10, Key .Code = "A"}
        Dim p2 = New With {Key .Id = 10, Key .Code = "A"}

        Console.WriteLine(p1.GetHashCode() = p2.GetHashCode())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_anonymous_type_to_string_format() {
    let src = r#"
Module Program
    Sub Main()
        Dim item = New With {.X = 5, .Y = 10}
        Console.WriteLine(item.ToString().Contains("X = 5"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
