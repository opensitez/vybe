use super::helpers::run_vb;

macro_rules! vb_spec_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, vec![$($expected),*]);
        }
    };
}

vb_spec_case!(named_arguments_can_reorder_all_parameters_in_function_call, r#"
Module M
    Function Describe(name As String, prefix As String, suffix As String) As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Describe(suffix:="!", name:="Ada", prefix:="Hi"))
    End Sub
End Module
"#, ["Hi:Ada:!"]);

vb_spec_case!(named_arguments_can_mix_positional_and_named_arguments, r#"
Module M
    Function Describe(name As String, prefix As String, suffix As String) As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Describe("Bea", suffix:="?", prefix:="Hello"))
    End Sub
End Module
"#, ["Hello:Bea:?"]);

vb_spec_case!(named_arguments_can_override_only_middle_optional_parameter, r#"
Module M
    Function Describe(name As String, Optional prefix As String = "start", Optional suffix As String = ".") As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Describe("Cora", prefix:="override"))
    End Sub
End Module
"#, ["override:Cora:."]);

vb_spec_case!(named_arguments_can_override_only_final_optional_parameter, r#"
Module M
    Function Describe(name As String, Optional prefix As String = "start", Optional suffix As String = ".") As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Describe("Dax", suffix:="done"))
    End Sub
End Module
"#, ["start:Dax:done"]);

vb_spec_case!(named_arguments_can_skip_to_last_optional_when_middle_is_positional, r#"
Module M
    Function Describe(name As String, Optional prefix As String = "base", Optional suffix As String = ".") As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Describe("Eli", "north", suffix:="!"))
    End Sub
End Module
"#, ["north:Eli:!"]);

vb_spec_case!(named_arguments_can_call_sub_with_boolean_and_message, r#"
Module M
    Sub PrintLine(message As String, uppercase As Boolean)
        If uppercase Then
            Console.WriteLine(message & "!")
        Else
            Console.WriteLine(message)
        End If
    End Sub

    Sub Main()
        PrintLine(uppercase:=True, message:="flagged")
    End Sub
End Module
"#, ["flagged!"]);

vb_spec_case!(named_arguments_can_call_instance_method_out_of_order, r#"
Class Formatter
    Public Function JoinParts(left As String, middle As String, right As String) As String
        Return left & "-" & middle & "-" & right
    End Function
End Class

Module M
    Sub Main()
        Dim formatter As New Formatter()
        Console.WriteLine(formatter.JoinParts(right:="finish", left:="start", middle:="middle"))
    End Sub
End Module
"#, ["start-middle-finish"]);

vb_spec_case!(named_arguments_can_call_shared_method_out_of_order, r#"
Class Formatter
    Public Shared Function Wrap(value As String, prefix As String, suffix As String) As String
        Return prefix & value & suffix
    End Function
End Class

Module M
    Sub Main()
        Console.WriteLine(Formatter.Wrap(suffix:="]", value:="core", prefix:="["))
    End Sub
End Module
"#, ["[core]"]);

vb_spec_case!(named_arguments_can_call_constructor_with_reordered_parameters, r#"
Class Person
    Public Name As String
    Public Age As Integer

    Public Sub New(age As Integer, name As String)
        Me.Name = name
        Me.Age = age
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Person(name:="Gus", age:=41)
        Console.WriteLine(p.Name)
        Console.WriteLine(p.Age)
    End Sub
End Module
"#, ["Gus", "41"]);

vb_spec_case!(named_arguments_can_use_expression_values, r#"
Module M
    Function Compute(total As Integer, scale As Integer, offset As Integer) As Integer
        Return total * scale + offset
    End Function

    Sub Main()
        Console.WriteLine(Compute(offset:=3, total:=5 + 1, scale:=2))
    End Sub
End Module
"#, ["15"]);

vb_spec_case!(named_arguments_can_use_function_results_as_values, r#"
Module M
    Function Prefix() As String
        Return "pre"
    End Function

    Function Suffix() As String
        Return "post"
    End Function

    Function Describe(name As String, prefix As String, suffix As String) As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Describe(name:="Hope", suffix:=Suffix(), prefix:=Prefix()))
    End Sub
End Module
"#, ["pre:Hope:post"]);

vb_spec_case!(named_arguments_can_target_byref_parameter_by_name, r#"
Module M
    Sub SetValue(ByRef value As Integer, amount As Integer)
        value = amount
    End Sub

    Sub Main()
        Dim total As Integer = 0
        SetValue(amount:=9, value:=total)
        Console.WriteLine(total)
    End Sub
End Module
"#, ["9"]);

vb_spec_case!(named_arguments_can_target_property_setter_parameter_order, r#"
Class Counter
    Public Value As Integer

    Public Sub Apply(amount As Integer, repeatCount As Integer)
        For i As Integer = 1 To repeatCount
            Value = Value + amount
        Next
    End Sub
End Class

Module M
    Sub Main()
        Dim counter As New Counter()
        counter.Apply(repeatCount:=3, amount:=4)
        Console.WriteLine(counter.Value)
    End Sub
End Module
"#, ["12"]);

vb_spec_case!(named_arguments_can_be_used_twice_with_different_orderings, r#"
Module M
    Function Describe(name As String, prefix As String, suffix As String) As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Describe(prefix:="A", suffix:="B", name:="first"))
        Console.WriteLine(Describe(name:="second", prefix:="C", suffix:="D"))
    End Sub
End Module
"#, ["A:first:B", "C:second:D"]);

vb_spec_case!(named_arguments_can_work_with_strings_containing_spaces, r#"
Module M
    Function Describe(name As String, prefix As String, suffix As String) As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Describe(name:="Ivy Lane", prefix:="hello there", suffix:="good bye"))
    End Sub
End Module
"#, ["hello there:Ivy Lane:good bye"]);