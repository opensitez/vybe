use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Version Parsing, Comparison & Components
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_version_construction_components() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ver As New Version(1, 2, 3, 4)
        Console.WriteLine(ver.Major & "." & ver.Minor & "." & ver.Build & "." & ver.Revision)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.2.3.4"]);
}

#[test]
fn test_vb_version_construction_two_components() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ver As New Version(2, 5)
        Console.WriteLine(ver.Major & "." & ver.Minor & "|Build=" & (ver.Build = -1) & "|Rev=" & (ver.Revision = -1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2.5|Build=True|Rev=True"]);
}

#[test]
fn test_vb_version_parse_standard_string() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ver = Version.Parse("3.14.15.926")
        Console.WriteLine(ver.Major & "|" & ver.Minor & "|" & ver.Build & "|" & ver.Revision)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3|14|15|926"]);
}

#[test]
fn test_vb_version_try_parse_success_and_failure() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ver As Version
        Dim ok = Version.TryParse("1.0.0", ver)
        Dim fail = Version.TryParse("InvalidVersion", ver)
        Console.WriteLine(ok & ":" & ver.ToString() & "|" & fail)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:1.0.0|False"]);
}

#[test]
fn test_vb_version_comparison_operators() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim v1 = New Version(1, 0, 0)
        Dim v2 = New Version(1, 1, 0)
        Dim v3 = New Version(2, 0, 0)
        Console.WriteLine((v1 < v2) & "|" & (v2 < v3) & "|" & (v1 = New Version(1, 0, 0)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_version_compare_to_instance_method() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim v1 = New Version(2, 1)
        Dim v2 = New Version(2, 0)
        Console.WriteLine(v1.CompareTo(v2) > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_version_major_revision_minor_revision() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' MajorRevision is High 16 bits of Revision, MinorRevision is Low 16 bits
        Dim ver As New Version(1, 0, 100, &H00020003)
        Console.WriteLine(ver.MajorRevision & "|" & ver.MinorRevision)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|3"]);
}

#[test]
fn test_vb_version_to_string_field_count_overload() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ver As New Version(1, 2, 3, 4)
        Console.WriteLine(ver.ToString(2) & "|" & ver.ToString(3) & "|" & ver.ToString(4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.2|1.2.3|1.2.3.4"]);
}

#[test]
fn test_vb_version_equality_and_hashcode() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim v1 = New Version(1, 2, 3)
        Dim v2 = New Version(1, 2, 3)
        Console.WriteLine((v1 = v2) & "|" & (v1.GetHashCode() = v2.GetHashCode()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_version_clone_shallow_copy() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim original As New Version(4, 5, 6)
        Dim cloned As Version = CType(original.Clone(), Version)
        Console.WriteLine((original = cloned) & "|" & (Not Object.ReferenceEquals(original, cloned)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_version_array_sorting() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim versions As Version() = {
            New Version(2, 0),
            New Version(1, 5),
            New Version(2, 0, 1),
            New Version(1, 0)
        }
        Array.Sort(versions)
        For Each v In versions
            Console.WriteLine(v.ToString())
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.0", "1.5", "2.0", "2.0.1"]);
}

#[test]
fn test_vb_version_parse_negative_component_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Version.Parse("1.-2.3")
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentOutOfRangeException Caught"]);
}

#[test]
fn test_vb_version_parse_too_many_components_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Version.Parse("1.2.3.4.5")
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentException Caught"]);
}

#[test]
fn test_vb_version_linq_query_filtering() {
    let src = r#"
Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim list = {New Version(1, 0), New Version(2, 0), New Version(3, 0)}
        Dim minV2 = list.Where(Function(v) v >= New Version(2, 0))
        For Each v In minV2
            Console.WriteLine(v.ToString())
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2.0", "3.0"]);
}

#[test]
fn test_vb_version_dictionary_lookup_by_version_key() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Version, String)()
        dict(New Version(1, 0, 0)) = "Release 1.0"
        dict(New Version(2, 0, 0)) = "Release 2.0"
        Console.WriteLine(dict(New Version(1, 0, 0)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Release 1.0"]);
}

#[test]
fn test_vb_version_to_string_field_count_invalid_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ver As New Version(1, 2)
        Try
            ver.ToString(3) ' Asking for 3 fields when only 2 exist!
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException Caught on FieldCount Overflow")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentException Caught on FieldCount Overflow"]
    );
}

#[test]
fn test_vb_version_structural_type_of_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim obj As Object = New Version(1, 0)
        Console.WriteLine(TypeOf obj Is Version)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_version_three_component_construction() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ver As New Version(1, 2, 3)
        Console.WriteLine(ver.Major & "." & ver.Minor & "." & ver.Build & "|Rev=" & (ver.Revision = -1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.2.3|Rev=True"]);
}

#[test]
fn test_vb_version_zero_components_all_zero() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim vZero As New Version(0, 0, 0, 0)
        Console.WriteLine(vZero.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0.0.0.0"]);
}

#[test]
fn test_vb_version_parse_single_component_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Version.Parse("1")
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException Caught on Single Component Version")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentException Caught on Single Component Version"]
    );
}
