use super::helpers::run_vb;

#[test]
fn extension_method_array() {
    let out = run_vb(
        r#"
Imports System.Runtime.CompilerServices

Module Extensions
    <Extension()>
    Public Function SumFirstTwo(arr() As Integer) As Integer
        Return arr(0) + arr(1)
    End Function
End Module

Module M
    Sub Main()
        Dim nums() As Integer = {5, 10, 15}
        Console.WriteLine(nums.SumFirstTwo())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["15"]);
}
