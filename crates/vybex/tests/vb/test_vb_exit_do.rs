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

vb_case!(exit_do_breaks_once_increment_reaches_target, r#"
Module M
    Sub Main()
        Dim total As Integer = 0
        Do
            total = total + 1
            If total >= 3 Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
"#, ["3"]);

vb_case!(exit_do_can_overshoot_threshold_before_stopping, r#"
Module M
    Sub Main()
        Dim total As Integer = 0
        Do
            total = total + 4
            If total >= 7 Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
"#, ["8"]);

vb_case!(exit_do_can_start_from_negative_value, r#"
Module M
    Sub Main()
        Dim total As Integer = -2
        Do
            total = total + 3
            If total >= 4 Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
"#, ["4"]);

vb_case!(exit_do_skips_tail_work_after_break, r#"
Module M
    Sub Main()
        Dim total As Integer = 0
        Do
            total = total + 1
            If total >= 2 Then Exit Do
            total = total + 10
        Loop
        Console.WriteLine(total)
    End Sub
End Module
"#, ["12"]);

vb_case!(exit_do_can_use_separate_counter_and_total, r#"
Module M
    Sub Main()
        Dim count As Integer = 0
        Dim total As Integer = 0
        Do
            count = count + 1
            total = total + count
            If count = 4 Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
"#, ["10"]);

vb_case!(exit_do_can_break_from_boolean_flag, r#"
Module M
    Sub Main()
        Dim count As Integer = 0
        Dim shouldStop As Boolean = False
        Do
            count = count + 1
            shouldStop = count >= 3
            If shouldStop Then Exit Do
        Loop
        Console.WriteLine(count)
    End Sub
End Module
"#, ["3"]);

vb_case!(exit_do_can_append_text_before_breaking, r#"
Module M
    Sub Main()
        Dim text As String = ""
        Dim count As Integer = 0
        Do
            count = count + 1
            text = text & count
            If count = 3 Then Exit Do
        Loop
        Console.WriteLine(text)
    End Sub
End Module
"#, ["123"]);

vb_case!(exit_do_can_use_helper_function_for_break_decision, r#"
Module M
    Function ReachedLimit(value As Integer) As Boolean
        Return value >= 5
    End Function

    Sub Main()
        Dim total As Integer = 0
        Do
            total = total + 2
            If ReachedLimit(total) Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
"#, ["6"]);

vb_case!(exit_do_can_break_after_even_iteration, r#"
Module M
    Sub Main()
        Dim count As Integer = 0
        Do
            count = count + 1
            If count Mod 2 = 0 Then
                Exit Do
            End If
        Loop
        Console.WriteLine(count)
    End Sub
End Module
"#, ["2"]);

vb_case!(exit_do_leaves_accumulator_available_after_loop, r#"
Module M
    Sub Main()
        Dim total As Integer = 1
        Do
            total = total * 2
            If total >= 8 Then Exit Do
        Loop
        Console.WriteLine(total + 1)
    End Sub
End Module
"#, ["9"]);