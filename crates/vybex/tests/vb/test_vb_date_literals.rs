use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_date_literal(month: i32, day: i32, year: i32) {
    let src = format!(r#"
Module M
    Sub Main()
        Dim d As Date = #{month}/{day}/{year}#
        Console.WriteLine(CStr(d))
    End Sub
End Module
"#, month = month, day = day, year = year);
    assert_vb_output_owned(src, vec![format!("{}/{}/{}", month, day, year)]);
}

macro_rules! date_literal_cases {
    ($($name:ident => ($month:expr, $day:expr, $year:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_date_literal($month, $day, $year); })*
    };
}

date_literal_cases! {
    date_literals_001 => (1, 1, 2024),
    date_literals_002 => (1, 15, 2024),
    date_literals_003 => (2, 1, 2024),
    date_literals_004 => (2, 14, 2024),
    date_literals_005 => (3, 1, 2024),
    date_literals_006 => (3, 20, 2024),
    date_literals_007 => (4, 5, 2024),
    date_literals_008 => (4, 30, 2024),
    date_literals_009 => (5, 10, 2024),
    date_literals_010 => (5, 25, 2024),
}