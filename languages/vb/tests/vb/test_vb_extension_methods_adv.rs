use super::helpers::run_vb;

#[test]
fn extension_methods_adv() {
    let out = run_vb(
        r#"
Imports System.Runtime.CompilerServices

Module StringExtensions
    <Extension()>
    Public Function Whisper(ByVal str As String) As String
        Return str.ToLower() & "..."
    End Function
End Module

Module M
    Sub Main()
        Dim msg As String = "HELLO"
        Console.WriteLine(msg.Whisper())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["hello..."]);
}
