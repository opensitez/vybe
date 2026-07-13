use super::helpers::run_vb;

#[test]
fn type_characters_variable_declaration() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Type characters define the type without explicit 'As Type'
        Dim i% = 10     ' Integer
        Dim l& = 100    ' Long
        Dim d@ = 10.5D  ' Decimal
        Dim s! = 2.5!   ' Single
        Dim f# = 3.14#  ' Double
        Dim str$ = "VB" ' String
        
        Console.WriteLine(i.GetType().Name)
        Console.WriteLine(l.GetType().Name)
        Console.WriteLine(d.GetType().Name)
        Console.WriteLine(s.GetType().Name)
        Console.WriteLine(f.GetType().Name)
        Console.WriteLine(str.GetType().Name)
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["Int32", "Int64", "Decimal", "Single", "Double", "String"]
    );
}

#[test]
fn type_characters_function_return() {
    let out = run_vb(
        r#"
Module M
    ' Function name with type character determines return type
    Function GetName$()
        Return "Alice"
    End Function

    Function GetAge%()
        Return 30
    End Function

    Sub Main()
        Console.WriteLine(GetName())
        Console.WriteLine(GetAge())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "30"]);
}
