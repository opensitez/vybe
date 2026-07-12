use super::helpers::run_vb;

macro_rules! vb_shared_spec {
    ($name:ident, $reason:expr, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, vec![$($expected),*]);
        }
    };
}

vb_shared_spec!(
    shared_field_retains_integer_state_across_shared_calls,
    "VB Shared field state is not implemented correctly yet",
    r#"
Class Counter
    Public Shared Total As Integer = 0

    Public Shared Sub AddOne()
        Total = Total + 1
    End Sub
End Class

Module M
    Sub Main()
        Counter.AddOne()
        Counter.AddOne()
        Console.WriteLine(Counter.Total)
    End Sub
End Module
"#,
    ["2"]
);

vb_shared_spec!(
    shared_field_can_start_from_nonzero_seed,
    "VB Shared field state is not implemented correctly yet",
    r#"
Class Counter
    Public Shared Total As Integer = 10

    Public Shared Sub AddStep()
        Total = Total + 5
    End Sub
End Class

Module M
    Sub Main()
        Counter.AddStep()
        Console.WriteLine(Counter.Total)
    End Sub
End Module
"#,
    ["15"]
);

vb_shared_spec!(
    shared_field_is_shared_across_multiple_instances,
    "VB Shared field state is not implemented correctly yet",
    r#"
Class Counter
    Public Shared Total As Integer = 0

    Public Sub AddOne()
        Total = Total + 1
    End Sub
End Class

Module M
    Sub Main()
        Dim left As New Counter()
        Dim right As New Counter()
        left.AddOne()
        right.AddOne()
        Console.WriteLine(Counter.Total)
    End Sub
End Module
"#,
    ["2"]
);

vb_shared_spec!(
    shared_field_can_be_read_through_instance_reference,
    "VB Shared field state is not implemented correctly yet",
    r#"
Class Counter
    Public Shared Total As Integer = 3
End Class

Module M
    Sub Main()
        Dim counter As New Counter()
        Console.WriteLine(counter.Total)
    End Sub
End Module
"#,
    ["3"]
);

vb_shared_spec!(
    shared_field_can_store_string_state,
    "VB Shared field state is not implemented correctly yet",
    r#"
Class Registry
    Public Shared Text As String = "A"

    Public Shared Sub Append(value As String)
        Text = Text & value
    End Sub
End Class

Module M
    Sub Main()
        Registry.Append("B")
        Registry.Append("C")
        Console.WriteLine(Registry.Text)
    End Sub
End Module
"#,
    ["ABC"]
);

vb_shared_spec!(
    shared_field_updates_from_helper_method_are_visible_later,
    "VB Shared field state is not implemented correctly yet",
    r#"
Class Counter
    Public Shared Total As Integer = 1

    Public Shared Sub ApplyDouble()
        Total = Total * 2
    End Sub
End Class

Module M
    Sub Main()
        Counter.ApplyDouble()
        Counter.ApplyDouble()
        Console.WriteLine(Counter.Total)
    End Sub
End Module
"#,
    ["4"]
);

vb_shared_spec!(
    shared_fields_are_isolated_between_classes,
    "VB Shared field state is not implemented correctly yet",
    r#"
Class LeftCounter
    Public Shared Total As Integer = 1
End Class

Class RightCounter
    Public Shared Total As Integer = 9
End Class

Module M
    Sub Main()
        Console.WriteLine(LeftCounter.Total)
        Console.WriteLine(RightCounter.Total)
    End Sub
End Module
"#,
    ["1", "9"]
);

vb_shared_spec!(
    shared_field_can_be_mutated_in_loop,
    "VB Shared field state is not implemented correctly yet",
    r#"
Class Counter
    Public Shared Total As Integer = 0

    Public Shared Sub AddStep()
        Total = Total + 2
    End Sub
End Class

Module M
    Sub Main()
        For i As Integer = 1 To 3
            Counter.AddStep()
        Next
        Console.WriteLine(Counter.Total)
    End Sub
End Module
"#,
    ["6"]
);
