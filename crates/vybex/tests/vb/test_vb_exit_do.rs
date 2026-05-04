use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_exit_do(limit: i32, step: i32) {
    let src = format!(r#"
Module M
    Sub Main()
        Dim total As Integer = 0
        Do
            total = total + {step}
            If total >= {limit} Then Exit Do
        Loop
        Console.WriteLine(total)
    End Sub
End Module
"#, limit = limit, step = step);
    let mut total = 0;
    loop {
        total += step;
        if total >= limit {
            break;
        }
    }
    assert_vb_output_owned(src, vec![total.to_string()]);
}

macro_rules! exit_do_cases {
    ($($name:ident => ($limit:expr, $step:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_exit_do($limit, $step); })*
    };
}

exit_do_cases! {
    exit_do_001 => (3, 1),
    exit_do_002 => (4, 1),
    exit_do_003 => (5, 2),
    exit_do_004 => (6, 2),
    exit_do_005 => (7, 3),
    exit_do_006 => (8, 3),
    exit_do_007 => (9, 4),
    exit_do_008 => (10, 4),
    exit_do_009 => (11, 5),
    exit_do_010 => (12, 5),
    exit_do_011 => (13, 2),
    exit_do_012 => (14, 3),
    exit_do_013 => (15, 4),
    exit_do_014 => (16, 5),
    exit_do_015 => (17, 6),
    exit_do_016 => (18, 6),
    exit_do_017 => (19, 7),
    exit_do_018 => (20, 7),
    exit_do_019 => (21, 3),
    exit_do_020 => (22, 4),
    exit_do_021 => (23, 5),
    exit_do_022 => (24, 6),
    exit_do_023 => (25, 7),
    exit_do_024 => (26, 8),
    exit_do_025 => (27, 9),
    exit_do_026 => (28, 4),
    exit_do_027 => (29, 5),
    exit_do_028 => (30, 6),
    exit_do_029 => (31, 7),
    exit_do_030 => (32, 8),
}