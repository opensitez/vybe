use super::helpers::run_vb;

#[test]
fn struct_layoutkind() {
    let out = run_vb(
        r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Explicit)>
Structure UnionType
    <FieldOffset(0)> Public I As Integer
    <FieldOffset(0)> Public B1 As Byte
    <FieldOffset(1)> Public B2 As Byte
    <FieldOffset(2)> Public B3 As Byte
    <FieldOffset(3)> Public B4 As Byte
End Structure

Module M
    Sub Main()
        Dim u As New UnionType()
        u.I = &H12345678
        ' B1 will be the least significant byte on little endian systems (0x78 = 120)
        Console.WriteLine(u.B1)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["120"]);
}
