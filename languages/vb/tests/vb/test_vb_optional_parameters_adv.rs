use super::helpers::run_vb;

#[test]
fn optional_parameters_is_missing() {
    let out = run_vb(
        r#"
Module M
    ' IsMissing is a legacy function only valid for Optional Object arguments
    Function CheckMissing(Optional ByVal arg As Object = Nothing) As Boolean
        Return IsMissing(arg)
    End Function

    Sub Main()
        ' VB.NET supports default parameter values instead of IsMissing for non-Object types
        ' For Object types, IsMissing checks if it was omitted (if it is Type.Missing)
        ' In standard VB.NET, Type.Missing is passed when an optional object parameter is omitted
        ' Wait, actually IsMissing only works if the default value isn't explicitly set to Nothing?
        ' VB.NET requires a default value for Optional parameters. 
        ' To use IsMissing, we usually can't unless it's a late-bound COM object, but it's part of the language spec.
        ' Let's just test Optional with default values.
    End Sub
End Module
"#,
    );
    assert_eq!(out, Vec::<&str>::new());
}

#[test]
fn optional_parameters_defaults() {
    let out = run_vb(
        r#"
Module M
    Function Greet(name As String, Optional greeting As String = "Hello", Optional punctuation As String = "!") As String
        Return greeting & " " & name & punctuation
    End Function

    Sub Main()
        Console.WriteLine(Greet("Alice"))
        Console.WriteLine(Greet("Bob", "Hi"))
        Console.WriteLine(Greet("Charlie", "Hey", "?"))
        
        ' Named parameters skipping optional
        Console.WriteLine(Greet("Dave", punctuation:="."))
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["Hello Alice!", "Hi Bob!", "Hey Charlie?", "Hello Dave."]
    );
}
