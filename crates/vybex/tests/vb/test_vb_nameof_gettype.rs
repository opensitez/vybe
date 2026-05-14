use super::helpers::run_vb;

macro_rules! vb_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, super::helpers::dotnet_expected_lines(&[$($expected),*]));
        }
    };
}

vb_case!(nameof_returns_local_variable_name, r#"
Module M
    Sub Main()
        Dim total As Integer = 5
        Console.WriteLine(NameOf(total))
    End Sub
End Module
"#, ["total"]);

vb_case!(nameof_returns_function_name, r#"
Module M
    Function ComputeTotal() As Integer
        Return 10
    End Function

    Sub Main()
        Console.WriteLine(NameOf(ComputeTotal))
    End Sub
End Module
"#, ["ComputeTotal"]);

vb_case!(nameof_returns_module_method_name, r#"
Module M
    Sub Main()
        Console.WriteLine(NameOf(Main))
    End Sub
End Module
"#, ["Main"]);

vb_case!(nameof_returns_type_name_for_builtin_type, r#"
Module M
    Sub Main()
        Console.WriteLine(NameOf(Integer))
    End Sub
End Module
"#, ["Integer"]);

vb_case!(nameof_returns_parameter_name_inside_function, r#"
Module M
    Function ShowName(value As Integer) As String
        Return NameOf(value)
    End Function

    Sub Main()
        Console.WriteLine(ShowName(5))
    End Sub
End Module
"#, ["value"]);

vb_case!(gettype_returns_non_nothing_for_integer, r#"
Module M
    Sub Main()
        Dim t As Object = GetType(Integer)
        If IsNothing(t) Then
            Console.WriteLine("missing")
        Else
            Console.WriteLine("present")
        End If
    End Sub
End Module
"#, ["present"]);

vb_case!(gettype_returns_non_nothing_for_string, r#"
Module M
    Sub Main()
        Dim t As Object = GetType(String)
        If IsNothing(t) Then
            Console.WriteLine("missing")
        Else
            Console.WriteLine("present")
        End If
    End Sub
End Module
"#, ["present"]);

vb_case!(gettype_returns_non_nothing_for_boolean, r#"
Module M
    Sub Main()
        Dim t As Object = GetType(Boolean)
        If IsNothing(t) Then
            Console.WriteLine("missing")
        Else
            Console.WriteLine("present")
        End If
    End Sub
End Module
"#, ["present"]);

vb_case!(gettype_returns_non_nothing_for_object, r#"
Module M
    Sub Main()
        Dim t As Object = GetType(Object)
        If IsNothing(t) Then
            Console.WriteLine("missing")
        Else
            Console.WriteLine("present")
        End If
    End Sub
End Module
"#, ["present"]);

vb_case!(gettype_returns_non_nothing_for_custom_class, r#"
Module M
    Class Greeter
    End Class

    Sub Main()
        Dim t As Object = GetType(Greeter)
        If IsNothing(t) Then
            Console.WriteLine("missing")
        Else
            Console.WriteLine("present")
        End If
    End Sub
End Module
"#, ["present"]);