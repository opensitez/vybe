use super::helpers::run_vb;

macro_rules! vb_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, vec![$($expected),*]);
        }
    };
}

vb_case!(date_literal_supports_new_years_day, r#"
Module M
    Sub Main()
        Dim d As Date = #1/1/2024#
        Console.WriteLine(CStr(d))
    End Sub
End Module
"#, ["1/1/2024"]);

vb_case!(date_literal_supports_mid_month_value, r#"
Module M
    Sub Main()
        Dim d As Date = #1/15/2024#
        Console.WriteLine(CStr(d))
    End Sub
End Module
"#, ["1/15/2024"]);

vb_case!(date_literal_supports_leap_day, r#"
Module M
    Sub Main()
        Dim d As Date = #2/29/2024#
        Console.WriteLine(CStr(d))
    End Sub
End Module
"#, ["2/29/2024"]);

vb_case!(date_literal_supports_summer_holiday_date, r#"
Module M
    Sub Main()
        Dim d As Date = #7/4/2024#
        Console.WriteLine(CStr(d))
    End Sub
End Module
"#, ["7/4/2024"]);

vb_case!(date_literal_supports_end_of_month_value, r#"
Module M
    Sub Main()
        Dim d As Date = #4/30/2024#
        Console.WriteLine(CStr(d))
    End Sub
End Module
"#, ["4/30/2024"]);

vb_case!(date_literal_supports_end_of_year_value, r#"
Module M
    Sub Main()
        Dim d As Date = #12/31/2024#
        Console.WriteLine(CStr(d))
    End Sub
End Module
"#, ["12/31/2024"]);

#[test]
fn date_literal_can_include_time_of_day_spec() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim d As Date = #5/14/2024 3:45 PM#
        Console.WriteLine(CStr(d))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["5/14/2024 3:45 PM"]);
}

#[test]
fn date_literal_can_include_seconds_precision() {
    let out = run_vb(r#"
Module M
    Sub Main()
        Dim d As Date = #5/14/2024 11:59:58 PM#
        Console.WriteLine(CStr(d))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["5/14/2024 11:59:58 PM"]);
}