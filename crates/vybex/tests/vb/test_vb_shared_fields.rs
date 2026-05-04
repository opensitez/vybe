use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_shared_fields(start: i32, step: i32, calls: i32) {
    let src = format!(r#"
Class Counter
    Public Shared Total As Integer = {start}

    Public Shared Sub AddStep()
        Total = Total + {step}
    End Sub
End Class

Module M
    Sub Main()
        For i As Integer = 1 To {calls}
            Counter.AddStep()
        Next
        Console.WriteLine(Counter.Total)
    End Sub
End Module
"#, start = start, step = step, calls = calls);
    assert_vb_output_owned(src, vec![(start + step * calls).to_string()]);
}

macro_rules! shared_field_cases {
    ($($name:ident => ($start:expr, $step:expr, $calls:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_shared_fields($start, $step, $calls); })*
    };
}

shared_field_cases! {
    shared_fields_001 => (0, 1, 1),
    shared_fields_002 => (5, 2, 2),
    shared_fields_003 => (10, 3, 3),
    shared_fields_004 => (20, 4, 1),
    shared_fields_005 => (25, 5, 2),
    shared_fields_006 => (30, 1, 5),
    shared_fields_007 => (40, 2, 4),
    shared_fields_008 => (50, 3, 2),
    shared_fields_009 => (60, 4, 3),
    shared_fields_010 => (70, 5, 1),
}