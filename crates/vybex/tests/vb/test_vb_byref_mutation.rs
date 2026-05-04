use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_byref_mutation(start: i32, delta: i32, calls: i32) {
    let src = format!(r#"
Module M
    Sub Bump(ByRef value As Integer)
        value = value + {delta}
    End Sub

    Sub Main()
        Dim x As Integer = {start}
        For i As Integer = 1 To {calls}
            Bump(x)
        Next
        Console.WriteLine(x)
    End Sub
End Module
"#, start = start, delta = delta, calls = calls);
    assert_vb_output_owned(src, vec![(start + delta * calls).to_string()]);
}

macro_rules! byref_mutation_cases {
    ($($name:ident => ($start:expr, $delta:expr, $calls:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_byref_mutation($start, $delta, $calls); })*
    };
}

byref_mutation_cases! {
    byref_mutation_001 => (-5, 1, 1),
    byref_mutation_002 => (0, 1, 1),
    byref_mutation_003 => (2, 3, 1),
    byref_mutation_004 => (4, 2, 2),
    byref_mutation_005 => (10, -1, 3),
    byref_mutation_006 => (7, 5, 2),
    byref_mutation_007 => (9, 0, 4),
    byref_mutation_008 => (12, 1, 5),
    byref_mutation_009 => (20, -2, 4),
    byref_mutation_010 => (3, 4, 3),
    byref_mutation_011 => (15, 2, 1),
    byref_mutation_012 => (18, 3, 2),
    byref_mutation_013 => (21, -3, 3),
    byref_mutation_014 => (24, 4, 1),
    byref_mutation_015 => (27, 1, 2),
    byref_mutation_016 => (30, 2, 3),
    byref_mutation_017 => (33, -1, 4),
    byref_mutation_018 => (36, 5, 1),
    byref_mutation_019 => (39, 2, 5),
    byref_mutation_020 => (42, -2, 2),
    byref_mutation_021 => (45, 3, 3),
    byref_mutation_022 => (48, 4, 2),
    byref_mutation_023 => (51, -3, 1),
    byref_mutation_024 => (54, 1, 4),
    byref_mutation_025 => (57, 2, 2),
    byref_mutation_026 => (60, 3, 1),
    byref_mutation_027 => (63, 0, 3),
    byref_mutation_028 => (66, -1, 2),
    byref_mutation_029 => (69, 5, 2),
    byref_mutation_030 => (72, 2, 4),
    byref_mutation_031 => (75, -2, 3),
    byref_mutation_032 => (78, 1, 1),
    byref_mutation_033 => (81, 4, 3),
    byref_mutation_034 => (84, 2, 2),
    byref_mutation_035 => (87, -1, 5),
    byref_mutation_036 => (90, 3, 2),
    byref_mutation_037 => (93, 1, 3),
    byref_mutation_038 => (96, -4, 2),
    byref_mutation_039 => (99, 5, 1),
    byref_mutation_040 => (102, 2, 5),
}