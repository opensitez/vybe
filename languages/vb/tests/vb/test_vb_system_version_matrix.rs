use super::helpers::run_vb;

#[test]
fn version_parse_full_text() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim version As Version = Version.Parse("2.3.4.5")
        Console.WriteLine(version.Major)
        Console.WriteLine(version.Minor)
        Console.WriteLine(version.Build)
        Console.WriteLine(version.Revision)
        Console.WriteLine(version.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "3", "4", "5", "2.3.4.5"]);
}

#[test]
fn version_parse_without_revision_keeps_build() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim version As Version = Version.Parse("5.6.7")
        Console.WriteLine(version.Major)
        Console.WriteLine(version.Minor)
        Console.WriteLine(version.Build)
        Console.WriteLine(version.Revision)
        Console.WriteLine(version.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "6", "7", "-1", "5.6.7"]);
}

#[test]
fn version_parse_short_form_keeps_missing_parts() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim version As Version = Version.Parse("7.8")
        Console.WriteLine(version.Major)
        Console.WriteLine(version.Minor)
        Console.WriteLine(version.Build)
        Console.WriteLine(version.Revision)
        Console.WriteLine(version.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["7", "8", "-1", "-1", "7.8"]);
}

#[test]
fn version_constructor_two_parts_uses_missing_build_and_revision_markers() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim version As New Version(9, 10)
        Console.WriteLine(version.Major)
        Console.WriteLine(version.Minor)
        Console.WriteLine(version.Build)
        Console.WriteLine(version.Revision)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["9", "10", "-1", "-1"]);
}

#[test]
fn version_constructor_four_parts_populates_all_fields() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim version As New Version(1, 2, 3, 4)
        Console.WriteLine(version.Major)
        Console.WriteLine(version.Minor)
        Console.WriteLine(version.Build)
        Console.WriteLine(version.Revision)
        Console.WriteLine(version.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "2", "3", "4", "1.2.3.4"]);
}

#[test]
fn version_new_defaults_to_zero_major_minor_shape() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim version As New Version()
        Console.WriteLine(version.Major)
        Console.WriteLine(version.Minor)
        Console.WriteLine(version.Build)
        Console.WriteLine(version.Revision)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0", "0", "-1", "-1"]);
}

#[test]
fn version_comparison_reflects_semantic_order() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim a As Version = Version.Parse("1.2.3")
        Dim b As Version = Version.Parse("1.2.4")
        Console.WriteLine(a.CompareTo(b) < 0)
        Console.WriteLine(b.CompareTo(a) > 0)
        Console.WriteLine(a.CompareTo(a))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "0"]);
}

#[test]
fn version_equals_ignores_instance_identity() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim left As New Version(3, 4, 5, 6)
        Dim right As New Version(3, 4, 5, 6)
        Dim parsed As Version = Version.Parse(left.ToString())

        Console.WriteLine(left.Equals(right))
        Console.WriteLine(left.Equals(parsed))
        Console.WriteLine(parsed.ToString() = left.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn version_roundtrip_preserves_parse_and_format() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim source As Version = Version.Parse("10.20.30.40")
        Dim roundTrip As Version = Version.Parse(source.ToString())
        Console.WriteLine(source.Major = roundTrip.Major)
        Console.WriteLine(source.Minor = roundTrip.Minor)
        Console.WriteLine(source.Build = roundTrip.Build)
        Console.WriteLine(source.Revision = roundTrip.Revision)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True", "True"]);
}
