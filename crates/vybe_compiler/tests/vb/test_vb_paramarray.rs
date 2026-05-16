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

vb_case!(paramarray_accepts_no_values, r#"
Module M
    Function CountValues(ParamArray values() As Integer) As Integer
        Return values.Length
    End Function

    Sub Main()
        Console.WriteLine(CountValues())
    End Sub
End Module
"#, ["0"]);

vb_case!(paramarray_sums_single_value, r#"
Module M
    Function SumAll(ParamArray values() As Integer) As Integer
        Dim total As Integer = 0
        For Each value As Integer In values
            total = total + value
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(SumAll(7))
    End Sub
End Module
"#, ["7"]);

vb_case!(paramarray_sums_many_values, r#"
Module M
    Function SumAll(ParamArray values() As Integer) As Integer
        Dim total As Integer = 0
        For Each value As Integer In values
            total = total + value
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(SumAll(1, 2, 3, 4, 5))
    End Sub
End Module
"#, ["15"]);

vb_case!(paramarray_handles_negative_numbers, r#"
Module M
    Function SumAll(ParamArray values() As Integer) As Integer
        Dim total As Integer = 0
        For Each value As Integer In values
            total = total + value
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(SumAll(10, -3, -2))
    End Sub
End Module
"#, ["5"]);

vb_case!(paramarray_handles_zero_values_without_special_cases, r#"
Module M
    Function SumAll(ParamArray values() As Integer) As Integer
        Dim total As Integer = 0
        For Each value As Integer In values
            total = total + value
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(SumAll(0, 0, 0))
    End Sub
End Module
"#, ["0"]);

vb_case!(paramarray_preserves_string_order_when_joining, r#"
Module M
    Function JoinAll(ParamArray values() As String) As String
        Dim result As String = ""
        For i As Integer = 0 To values.Length - 1
            If i > 0 Then
                result = result & ","
            End If
            result = result & values(i)
        Next
        Return result
    End Function

    Sub Main()
        Console.WriteLine(JoinAll("red", "green", "blue"))
    End Sub
End Module
"#, ["red,green,blue"]);

vb_case!(paramarray_can_count_arguments, r#"
Module M
    Function CountValues(ParamArray values() As Integer) As Integer
        Return values.Length
    End Function

    Sub Main()
        Console.WriteLine(CountValues(2, 4, 6, 8))
    End Sub
End Module
"#, ["4"]);

vb_case!(paramarray_can_follow_required_prefix_argument, r#"
Module M
    Function SumWithOffset(offset As Integer, ParamArray values() As Integer) As Integer
        Dim total As Integer = offset
        For Each value As Integer In values
            total = total + value
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(SumWithOffset(10, 1, 2, 3))
    End Sub
End Module
"#, ["16"]);

vb_case!(paramarray_can_be_used_in_shared_methods, r#"
Class MathBox
    Public Shared Function MultiplyAll(ParamArray values() As Integer) As Integer
        Dim total As Integer = 1
        For Each value As Integer In values
            total = total * value
        Next
        Return total
    End Function
End Class

Module M
    Sub Main()
        Console.WriteLine(MathBox.MultiplyAll(2, 3, 4))
    End Sub
End Module
"#, ["24"]);

vb_case!(paramarray_can_be_used_in_instance_methods, r#"
Class TextWriter
    Public Function JoinWithBar(ParamArray values() As String) As String
        Dim result As String = ""
        For i As Integer = 0 To values.Length - 1
            If i > 0 Then
                result = result & "|"
            End If
            result = result & values(i)
        Next
        Return result
    End Function
End Class

Module M
    Sub Main()
        Dim writer As New TextWriter()
        Console.WriteLine(writer.JoinWithBar("a", "b", "c"))
    End Sub
End Module
"#, ["a|b|c"]);

vb_case!(paramarray_can_compute_maximum_value, r#"
Module M
    Function MaxValue(ParamArray values() As Integer) As Integer
        Dim current As Integer = values(0)
        For Each value As Integer In values
            If value > current Then
                current = value
            End If
        Next
        Return current
    End Function

    Sub Main()
        Console.WriteLine(MaxValue(4, 9, 1, 7))
    End Sub
End Module
"#, ["9"]);

vb_case!(paramarray_can_return_first_and_last_values, r#"
Module M
    Function EdgeValues(ParamArray values() As Integer) As String
        Return values(0) & ":" & values(values.Length - 1)
    End Function

    Sub Main()
        Console.WriteLine(EdgeValues(5, 6, 7, 8))
    End Sub
End Module
"#, ["5:8"]);

vb_case!(paramarray_accepts_expression_arguments, r#"
Module M
    Function SumAll(ParamArray values() As Integer) As Integer
        Dim total As Integer = 0
        For Each value As Integer In values
            total = total + value
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(SumAll(1 + 2, 3 * 2, 10 - 4))
    End Sub
End Module
"#, ["15"]);

vb_case!(paramarray_can_filter_even_numbers, r#"
Module M
    Function CountEven(ParamArray values() As Integer) As Integer
        Dim total As Integer = 0
        For Each value As Integer In values
            If value Mod 2 = 0 Then
                total = total + 1
            End If
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(CountEven(1, 2, 4, 7, 8))
    End Sub
End Module
"#, ["3"]);

vb_case!(paramarray_skips_loops_for_empty_input, r#"
Module M
    Function JoinAll(ParamArray values() As String) As String
        Dim result As String = "empty"
        For Each value As String In values
            result = result & value
        Next
        Return result
    End Function

    Sub Main()
        Console.WriteLine(JoinAll())
    End Sub
End Module
"#, ["empty"]);