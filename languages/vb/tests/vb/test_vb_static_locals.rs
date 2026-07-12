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

vb_case!(
    static_local_counter_retains_value_across_calls,
    r#"
Module M
    Function Counter() As Integer
        Static total As Integer = 0
        total = total + 1
        Return total
    End Function

    Sub Main()
        Console.WriteLine(Counter())
        Console.WriteLine(Counter())
        Console.WriteLine(Counter())
    End Sub
End Module
"#,
    ["1", "2", "3"]
);

vb_case!(
    static_local_counter_handles_negative_step,
    r#"
Module M
    Function Counter() As Integer
        Static total As Integer = 10
        total = total - 2
        Return total
    End Function

    Sub Main()
        Console.WriteLine(Counter())
        Console.WriteLine(Counter())
    End Sub
End Module
"#,
    ["8", "6"]
);

vb_case!(
    static_local_counter_handles_zero_step,
    r#"
Module M
    Function Counter() As Integer
        Static total As Integer = 5
        total = total + 0
        Return total
    End Function

    Sub Main()
        Console.WriteLine(Counter())
        Console.WriteLine(Counter())
    End Sub
End Module
"#,
    ["5", "5"]
);

vb_case!(
    static_local_string_accumulates_suffixes,
    r#"
Module M
    Function AppendNext(piece As String) As String
        Static text As String = ""
        text = text & piece
        Return text
    End Function

    Sub Main()
        Console.WriteLine(AppendNext("A"))
        Console.WriteLine(AppendNext("B"))
        Console.WriteLine(AppendNext("C"))
    End Sub
End Module
"#,
    ["A", "AB", "ABC"]
);

vb_case!(
    static_local_boolean_toggle_flips_each_call,
    r#"
Module M
    Function Toggle() As Boolean
        Static current As Boolean = False
        current = Not current
        Return current
    End Function

    Sub Main()
        Console.WriteLine(Toggle())
        Console.WriteLine(Toggle())
        Console.WriteLine(Toggle())
    End Sub
End Module
"#,
    ["true", "false", "true"]
);

vb_case!(
    static_local_is_isolated_per_function,
    r#"
Module M
    Function LeftCounter() As Integer
        Static total As Integer = 0
        total = total + 1
        Return total
    End Function

    Function RightCounter() As Integer
        Static total As Integer = 10
        total = total + 5
        Return total
    End Function

    Sub Main()
        Console.WriteLine(LeftCounter())
        Console.WriteLine(RightCounter())
        Console.WriteLine(LeftCounter())
        Console.WriteLine(RightCounter())
    End Sub
End Module
"#,
    ["1", "15", "2", "20"]
);

vb_case!(
    static_local_multiple_values_can_coexist_in_one_function,
    r#"
Module M
    Function Snapshot() As String
        Static count As Integer = 0
        Static text As String = "seed"
        count = count + 1
        text = text & count
        Return count & ":" & text
    End Function

    Sub Main()
        Console.WriteLine(Snapshot())
        Console.WriteLine(Snapshot())
    End Sub
End Module
"#,
    ["1:seed1", "2:seed12"]
);

vb_case!(
    static_local_in_sub_tracks_invocation_count,
    r#"
Module M
    Sub Report()
        Static callCount As Integer = 0
        callCount = callCount + 1
        Console.WriteLine(callCount)
    End Sub

    Sub Main()
        Report()
        Report()
        Report()
    End Sub
End Module
"#,
    ["1", "2", "3"]
);

vb_case!(
    static_local_can_remember_last_argument,
    r#"
Module M
    Function Remember(value As Integer) As String
        Static previous As Integer = -1
        Dim result As String = previous & "->" & value
        previous = value
        Return result
    End Function

    Sub Main()
        Console.WriteLine(Remember(4))
        Console.WriteLine(Remember(9))
        Console.WriteLine(Remember(2))
    End Sub
End Module
"#,
    ["-1->4", "4->9", "9->2"]
);

vb_case!(
    static_local_can_accumulate_string_lengths,
    r#"
Module M
    Function CountChars(text As String) As Integer
        Static total As Integer = 0
        total = total + Len(text)
        Return total
    End Function

    Sub Main()
        Console.WriteLine(CountChars("hi"))
        Console.WriteLine(CountChars("there"))
    End Sub
End Module
"#,
    ["2", "7"]
);

vb_case!(
    static_local_can_drive_loop_guard,
    r#"
Module M
    Function NextValue() As Integer
        Static total As Integer = 0
        total = total + 2
        Return total
    End Function

    Sub Main()
        Do While NextValue() < 7
            Console.WriteLine("loop")
        Loop
        Console.WriteLine(NextValue())
    End Sub
End Module
"#,
    ["loop", "loop", "loop", "10"]
);

vb_case!(
    static_local_two_string_functions_do_not_share_state,
    r#"
Module M
    Function LeftText(value As String) As String
        Static text As String = "L"
        text = text & value
        Return text
    End Function

    Function RightText(value As String) As String
        Static text As String = "R"
        text = text & value
        Return text
    End Function

    Sub Main()
        Console.WriteLine(LeftText("a"))
        Console.WriteLine(RightText("b"))
        Console.WriteLine(LeftText("c"))
        Console.WriteLine(RightText("d"))
    End Sub
End Module
"#,
    ["La", "Rb", "Lac", "Rbd"]
);

vb_case!(
    static_local_can_preserve_branch_updates,
    r#"
Module M
    Function Track(flag As Boolean) As Integer
        Static total As Integer = 0
        If flag Then
            total = total + 10
        Else
            total = total + 1
        End If
        Return total
    End Function

    Sub Main()
        Console.WriteLine(Track(False))
        Console.WriteLine(Track(True))
        Console.WriteLine(Track(False))
    End Sub
End Module
"#,
    ["1", "11", "12"]
);

vb_case!(
    static_local_can_start_from_negative_seed,
    r#"
Module M
    Function Counter() As Integer
        Static total As Integer = -5
        total = total + 3
        Return total
    End Function

    Sub Main()
        Console.WriteLine(Counter())
        Console.WriteLine(Counter())
    End Sub
End Module
"#,
    ["-2", "1"]
);

vb_case!(
    static_local_can_store_boolean_and_count_together,
    r#"
Module M
    Function Snapshot() As String
        Static flag As Boolean = False
        Static count As Integer = 0
        flag = Not flag
        If flag Then
            count = count + 1
        End If
        Return flag & ":" & count
    End Function

    Sub Main()
        Console.WriteLine(Snapshot())
        Console.WriteLine(Snapshot())
        Console.WriteLine(Snapshot())
    End Sub
End Module
"#,
    ["true:1", "false:1", "true:2"]
);

vb_case!(
    static_local_can_be_used_from_instance_method,
    r#"
Class Worker
    Public Function NextId() As Integer
        Static current As Integer = 100
        current = current + 1
        Return current
    End Function
End Class

Module M
    Sub Main()
        Dim worker As New Worker()
        Console.WriteLine(worker.NextId())
        Console.WriteLine(worker.NextId())
    End Sub
End Module
"#,
    ["101", "102"]
);

vb_case!(
    static_local_can_be_used_from_shared_method,
    r#"
Class Worker
    Public Shared Function NextBatch() As Integer
        Static batch As Integer = 1
        batch = batch * 2
        Return batch
    End Function
End Class

Module M
    Sub Main()
        Console.WriteLine(Worker.NextBatch())
        Console.WriteLine(Worker.NextBatch())
        Console.WriteLine(Worker.NextBatch())
    End Sub
End Module
"#,
    ["2", "4", "8"]
);

vb_case!(
    static_local_can_accumulate_with_varying_arguments,
    r#"
Module M
    Function AddValue(value As Integer) As Integer
        Static total As Integer = 0
        total = total + value
        Return total
    End Function

    Sub Main()
        Console.WriteLine(AddValue(3))
        Console.WriteLine(AddValue(7))
        Console.WriteLine(AddValue(-2))
    End Sub
End Module
"#,
    ["3", "10", "8"]
);
