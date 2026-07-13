use super::helpers::run_vb;

#[test]
fn trycast_linq() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim objs() As Object = {"A", 1, "B", 2}
        
        Dim strings = From o In objs
                      Let s = TryCast(o, String)
                      Where s IsNot Nothing
                      Select s
                      
        For Each s In strings
            Console.WriteLine(s)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["A", "B"]);
}
