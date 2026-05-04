use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_static_local_numeric(start: i32, step: i32, calls: i32) {
    let src = format!(r#"
Module M
    Function Counter() As Integer
        Static total As Integer = {start}
        total = total + {step}
        Return total
    End Function

    Sub Main()
        For i As Integer = 1 To {calls}
            Console.WriteLine(Counter())
        Next
    End Sub
End Module
"#, start = start, step = step, calls = calls);
    let expected: Vec<String> = (1..=calls).map(|index| (start + step * index).to_string()).collect();
    assert_vb_output_owned(src, expected);
}

macro_rules! static_local_cases {
    ($($name:ident => ($start:expr, $step:expr, $calls:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_static_local_numeric($start, $step, $calls); })*
    };
}

static_local_cases! {
    static_locals_001 => (0, 1, 2),
    static_locals_002 => (5, 1, 3),
    static_locals_003 => (10, 2, 2),
    static_locals_004 => (3, 4, 3),
    static_locals_005 => (8, 5, 2),
    static_locals_006 => (12, -1, 3),
    static_locals_007 => (20, 2, 4),
    static_locals_008 => (1, 3, 3),
    static_locals_009 => (7, 7, 2),
    static_locals_010 => (15, 5, 3),
    static_locals_011 => (2, 6, 2),
    static_locals_012 => (9, 4, 4),
    static_locals_013 => (30, 1, 5),
    static_locals_014 => (11, 2, 3),
    static_locals_015 => (14, 3, 2),
    static_locals_016 => (18, 4, 3),
    static_locals_017 => (25, -2, 3),
    static_locals_018 => (6, 8, 2),
    static_locals_019 => (13, 5, 4),
    static_locals_020 => (21, 2, 3),
    static_locals_021 => (4, 9, 2),
    static_locals_022 => (16, 1, 4),
    static_locals_023 => (22, 3, 3),
    static_locals_024 => (28, 2, 2),
    static_locals_025 => (32, 5, 3),
    static_locals_026 => (40, -3, 2),
    static_locals_027 => (50, 4, 4),
    static_locals_028 => (60, 1, 3),
    static_locals_029 => (70, 2, 2),
    static_locals_030 => (80, 3, 3),
    static_locals_031 => (90, 4, 2),
    static_locals_032 => (100, 5, 3),
    static_locals_033 => (110, 1, 4),
    static_locals_034 => (120, 2, 3),
    static_locals_035 => (130, 3, 2),
    static_locals_036 => (140, 4, 3),
    static_locals_037 => (150, 5, 2),
    static_locals_038 => (160, -1, 4),
    static_locals_039 => (170, 2, 3),
    static_locals_040 => (180, 3, 2),
}