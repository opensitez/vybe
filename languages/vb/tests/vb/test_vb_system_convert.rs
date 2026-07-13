use super::helpers::run_vb;

#[test]
fn system_convert_base64() {
    let out = run_vb(
        r#"
Imports System
Imports System.Text

Module M
    Sub Main()
        Dim original As String = "VB.NET Rocks!"
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(original)
        
        Dim base64 As String = Convert.ToBase64String(bytes)
        Console.WriteLine(base64 IsNot Nothing)
        
        Dim decodedBytes As Byte() = Convert.FromBase64String(base64)
        Dim decoded As String = Encoding.UTF8.GetString(decodedBytes)
        
        Console.WriteLine(decoded = original)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn system_convert_primitives() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim s As String = "123"
        Dim i As Integer = Convert.ToInt32(s)
        Console.WriteLine(i)
        
        Dim b As Boolean = Convert.ToBoolean(1)
        Console.WriteLine(b)
        
        Dim d As Double = Convert.ToDouble("3.14", Globalization.CultureInfo.InvariantCulture)
        Console.WriteLine(d)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["123", "True", "3.14"]);
}
