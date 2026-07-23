use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Guid Generation, Formats & Byte Arrays
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_guid_new_guid_uniqueness() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim g1 = Guid.NewGuid()
        Dim g2 = Guid.NewGuid()
        Console.WriteLine((g1 <> g2) & "|" & (g1 <> Guid.Empty))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_guid_empty_singleton() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim gEmpty = Guid.Empty
        Console.WriteLine(gEmpty.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["00000000-0000-0000-0000-000000000000"]);
}

#[test]
fn test_vb_guid_parse_standard_hyphenated() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim str = "d3b07384-d113-40a6-a719-88125d4699d5"
        Dim g = Guid.Parse(str)
        Console.WriteLine(g.ToString("D"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["d3b07384-d113-40a6-a719-88125d4699d5"]);
}

#[test]
fn test_vb_guid_to_string_format_specifiers_n_d_b_p_x() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim str = "d3b07384-d113-40a6-a719-88125d4699d5"
        Dim g = Guid.Parse(str)
        Dim formatN = g.ToString("N") ' No hyphens (32 hex digits)
        Dim formatB = g.ToString("B") ' Enclosed in braces { ... }
        Dim formatP = g.ToString("P") ' Enclosed in parentheses ( ... )

        Console.WriteLine(formatN.Length & "|" & formatB.StartsWith("{") & "|" & formatP.StartsWith("("))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["32|True|True"]);
}

#[test]
fn test_vb_guid_parse_exact_format_n() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim raw = "d3b07384d11340a6a71988125d4699d5"
        Dim g = Guid.ParseExact(raw, "N")
        Console.WriteLine(g.ToString("N"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["d3b07384d11340a6a71988125d4699d5"]);
}

#[test]
fn test_vb_guid_parse_exact_format_b_braces() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim raw = "{d3b07384-d113-40a6-a719-88125d4699d5}"
        Dim g = Guid.ParseExact(raw, "B")
        Console.WriteLine(g.ToString("B"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["{d3b07384-d113-40a6-a719-88125d4699d5}"]);
}

#[test]
fn test_vb_guid_try_parse_success_and_failure() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim g As Guid
        Dim ok = Guid.TryParse("d3b07384-d113-40a6-a719-88125d4699d5", g)
        Dim fail = Guid.TryParse("InvalidGuidString", g)
        Console.WriteLine(ok & "|" & fail)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_guid_to_byte_array_roundtrip() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim orig = Guid.NewGuid()
        Dim bytes = orig.ToByteArray()
        Dim restored As New Guid(bytes)
        Console.WriteLine((orig = restored) & "|" & (bytes.Length = 16))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_guid_constructor_byte_array() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes(15) As Byte
        bytes(0) = 1
        bytes(15) = 255
        Dim g As New Guid(bytes)
        Console.WriteLine(g <> Guid.Empty)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_guid_constructor_integer_short_byte_components() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' Guid(a As Integer, b As Short, c As Short, d As Byte, e As Byte, f As Byte, g As Byte, h As Byte, i As Byte, j As Byte, k As Byte)
        Dim g As New Guid(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
        Console.WriteLine(g.ToString("N").StartsWith("00000001"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_guid_equality_operator() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim str = "d3b07384-d113-40a6-a719-88125d4699d5"
        Dim g1 = Guid.Parse(str)
        Dim g2 = Guid.Parse(str)
        Dim g3 = Guid.NewGuid()
        Console.WriteLine((g1 = g2) & "|" & (g1 <> g3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_guid_compare_to_ordering() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim g1 = Guid.Empty
        Dim g2 = Guid.NewGuid()
        Console.WriteLine(g1.CompareTo(g2) < 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_guid_hashset_deduplication() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim g = Guid.NewGuid()
        Dim set As New HashSet(Of Guid)()
        set.Add(g)
        set.Add(g)
        Console.WriteLine(set.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_guid_dictionary_key_lookup() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim key = Guid.NewGuid()
        Dim dict As New Dictionary(Of Guid, String)()
        dict(key) = "FoundValue"
        Console.WriteLine(dict(key))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FoundValue"]);
}

#[test]
fn test_vb_guid_parse_invalid_length_throws_format_exception() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Guid.Parse("12345")
        Catch ex As FormatException
            Console.WriteLine("FormatException Caught on Short Guid")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FormatException Caught on Short Guid"]);
}

#[test]
fn test_vb_guid_try_parse_exact_validates_format_strictly() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim g As Guid
        ' Standard hyphenated GUID string passed to "N" (no hyphens) format expected!
        Dim ok = Guid.TryParseExact("d3b07384-d113-40a6-a719-88125d4699d5", "N", g)
        Console.WriteLine(ok)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_guid_format_x_structure_literal() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim g = Guid.Parse("d3b07384-d113-40a6-a719-88125d4699d5")
        Dim strX = g.ToString("X")
        Console.WriteLine(strX.StartsWith("{0x"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_guid_null_or_empty_constructor_byte_array_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim g As New Guid(New Byte(10) {}) ' Length 11 instead of 16
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException Caught on Invalid Byte Array Length")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentException Caught on Invalid Byte Array Length"]
    );
}

#[test]
fn test_vb_guid_structural_type_of_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim obj As Object = Guid.NewGuid()
        Console.WriteLine(TypeOf obj Is Guid)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_guid_to_string_uppercase() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim g = Guid.Parse("d3b07384-d113-40a6-a719-88125d4699d5")
        Dim upperStr = g.ToString("D").ToUpper()
        Console.WriteLine(upperStr.StartsWith("D3B07384"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
