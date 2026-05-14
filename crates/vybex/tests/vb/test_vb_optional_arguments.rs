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

vb_case!(optional_arguments_use_both_defaults_when_omitted, r#"
Module M
    Function Decorate(name As String, Optional prefix As String = "Hello", Optional suffix As String = "!") As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Decorate("Ada"))
    End Sub
End Module
"#, ["Hello:Ada:!"]);

vb_case!(optional_arguments_override_first_optional_only, r#"
Module M
    Function Decorate(name As String, Optional prefix As String = "Hello", Optional suffix As String = "!") As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Decorate("Bea", "Hi"))
    End Sub
End Module
"#, ["Hi:Bea:!"]);

vb_case!(optional_arguments_override_both_optional_values, r#"
Module M
    Function Decorate(name As String, Optional prefix As String = "Hello", Optional suffix As String = "!") As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Decorate("Cora", "Welcome", "?"))
    End Sub
End Module
"#, ["Welcome:Cora:?"]);

vb_case!(optional_arguments_support_integer_defaults, r#"
Module M
    Function AddBonus(value As Integer, Optional bonus As Integer = 5) As Integer
        Return value + bonus
    End Function

    Sub Main()
        Console.WriteLine(AddBonus(7))
    End Sub
End Module
"#, ["12"]);

vb_case!(optional_arguments_support_integer_override, r#"
Module M
    Function AddBonus(value As Integer, Optional bonus As Integer = 5) As Integer
        Return value + bonus
    End Function

    Sub Main()
        Console.WriteLine(AddBonus(7, 11))
    End Sub
End Module
"#, ["18"]);

vb_case!(optional_arguments_drive_boolean_default_branch, r#"
Module M
    Function Render(label As String, Optional loud As Boolean = False) As String
        If loud Then
            Return label & "!"
        End If
        Return label & "."
    End Function

    Sub Main()
        Console.WriteLine(Render("calm"))
    End Sub
End Module
"#, ["calm."]);

vb_case!(optional_arguments_can_override_boolean_branch, r#"
Module M
    Function Render(label As String, Optional loud As Boolean = False) As String
        If loud Then
            Return label & "!"
        End If
        Return label & "."
    End Function

    Sub Main()
        Console.WriteLine(Render("loud", True))
    End Sub
End Module
"#, ["loud!"]);

vb_case!(optional_arguments_work_in_shared_methods, r#"
Class Formatter
    Public Shared Function Wrap(value As String, Optional prefix As String = "[", Optional suffix As String = "]") As String
        Return prefix & value & suffix
    End Function
End Class

Module M
    Sub Main()
        Console.WriteLine(Formatter.Wrap("core"))
        Console.WriteLine(Formatter.Wrap("value", "<", ">"))
    End Sub
End Module
"#, ["[core]", "<value>"]);

vb_case!(optional_arguments_work_in_instance_methods, r#"
Class Greeter
    Public Function Build(name As String, Optional prefix As String = "Hi") As String
        Return prefix & " " & name
    End Function
End Class

Module M
    Sub Main()
        Dim greeter As New Greeter()
        Console.WriteLine(greeter.Build("Dana"))
        Console.WriteLine(greeter.Build("Eli", "Hello"))
    End Sub
End Module
"#, ["Hi Dana", "Hello Eli"]);

vb_case!(optional_arguments_drive_sub_side_effects, r#"
Module M
    Sub AppendLine(label As String, Optional suffix As String = ".")
        Console.WriteLine(label & suffix)
    End Sub

    Sub Main()
        AppendLine("first")
        AppendLine("second", "!")
    End Sub
End Module
"#, ["first.", "second!"]);

vb_case!(optional_arguments_can_chain_through_helper_functions, r#"
Module M
    Function Decorate(name As String, Optional prefix As String = "base", Optional suffix As String = ".") As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Function Outer(name As String, Optional prefix As String = "outer") As String
        Return Decorate(name, prefix)
    End Function

    Sub Main()
        Console.WriteLine(Outer("Faye"))
        Console.WriteLine(Outer("Gus", "inner"))
    End Sub
End Module
"#, ["outer:Faye:.", "inner:Gus:."]);

vb_case!(optional_arguments_allow_empty_string_defaults, r#"
Module M
    Function Wrap(value As String, Optional prefix As String = "", Optional suffix As String = "") As String
        Return prefix & value & suffix
    End Function

    Sub Main()
        Console.WriteLine(Wrap("plain"))
        Console.WriteLine(Wrap("tag", "<", ">"))
    End Sub
End Module
"#, ["plain", "<tag>"]);

vb_case!(optional_arguments_can_control_loop_iterations, r#"
Module M
    Function CountUp(Optional repeatCount As Integer = 3) As Integer
        Dim total As Integer = 0
        For i As Integer = 1 To repeatCount
            total = total + i
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(CountUp())
        Console.WriteLine(CountUp(4))
    End Sub
End Module
"#, ["6", "10"]);

vb_case!(optional_arguments_support_multiple_types_in_single_signature, r#"
Module M
    Function Describe(name As String, Optional level As Integer = 1, Optional suffix As String = "ok") As String
        Return name & ":" & level & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Describe("Hope"))
        Console.WriteLine(Describe("Ivy", 3, "done"))
    End Sub
End Module
"#, ["Hope:1:ok", "Ivy:3:done"]);

vb_case!(optional_arguments_can_use_second_default_after_first_override, r#"
Module M
    Function Build(name As String, Optional prefix As String = "start", Optional suffix As String = "end") As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Build("Jade", "custom"))
    End Sub
End Module
"#, ["custom:Jade:end"]);