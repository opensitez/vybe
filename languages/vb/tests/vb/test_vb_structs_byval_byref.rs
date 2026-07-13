use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Structs ByVal and ByRef
// ═══════════════════════════════════════════════════════════

#[test]
fn struct_byval_parameter() {
    let out = run_vb(
        r#"
Structure Coordinate
    Public Lat As Double
    Public Lon As Double
End Structure

Module M
    Sub ModifyCoord(ByVal c As Coordinate)
        c.Lat = 99.9
    End Sub

    Sub Main()
        Dim loc As Coordinate
        loc.Lat = 45.0
        ModifyCoord(loc)
        ' Should remain unchanged because it was passed by value (copied)
        Console.WriteLine(loc.Lat)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["45"]);
}

#[test]
fn struct_byref_parameter() {
    let out = run_vb(
        r#"
Structure Coordinate
    Public Lat As Double
    Public Lon As Double
End Structure

Module M
    Sub ModifyCoord(ByRef c As Coordinate)
        c.Lat = 99.9
    End Sub

    Sub Main()
        Dim loc As Coordinate
        loc.Lat = 45.0
        ModifyCoord(loc)
        ' Should change because it was passed by reference
        Console.WriteLine(loc.Lat)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["99.9"]);
}
