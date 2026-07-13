use super::helpers::run_vb;

// DateAdd tests
#[test] fn dateadd_days() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateAdd("d", 5, #1/1/2020#).Day): End Sub: End Module"#), vec!["6"]); }
#[test] fn dateadd_months() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateAdd("m", 2, #1/1/2020#).Month): End Sub: End Module"#), vec!["3"]); }
#[test] fn dateadd_years() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateAdd("yyyy", 10, #1/1/2020#).Year): End Sub: End Module"#), vec!["2030"]); }
#[test] fn dateadd_hours() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateAdd("h", 5, #12:00:00 PM#).Hour): End Sub: End Module"#), vec!["17"]); }
#[test] fn dateadd_minutes() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateAdd("n", 30, #12:00:00 PM#).Minute): End Sub: End Module"#), vec!["30"]); }
#[test] fn dateadd_seconds() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateAdd("s", 45, #12:00:00 PM#).Second): End Sub: End Module"#), vec!["45"]); }
#[test] fn dateadd_quarters() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateAdd("q", 1, #1/1/2020#).Month): End Sub: End Module"#), vec!["4"]); }
#[test] fn dateadd_weeks() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateAdd("ww", 2, #1/1/2020#).Day): End Sub: End Module"#), vec!["15"]); }
#[test] fn dateadd_weekdays() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateAdd("w", 5, #1/1/2020#).Day): End Sub: End Module"#), vec!["6"]); }
#[test] fn dateadd_dayofyear() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateAdd("y", 10, #1/1/2020#).Day): End Sub: End Module"#), vec!["11"]); }

// DateDiff tests
#[test] fn datediff_days() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateDiff("d", #1/1/2020#, #1/6/2020#)): End Sub: End Module"#), vec!["5"]); }
#[test] fn datediff_months() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateDiff("m", #1/1/2020#, #3/1/2020#)): End Sub: End Module"#), vec!["2"]); }
#[test] fn datediff_years() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateDiff("yyyy", #1/1/2020#, #1/1/2030#)): End Sub: End Module"#), vec!["10"]); }
#[test] fn datediff_hours() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateDiff("h", #12:00:00 PM#, #5:00:00 PM#)): End Sub: End Module"#), vec!["5"]); }
#[test] fn datediff_minutes() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateDiff("n", #12:00:00 PM#, #12:30:00 PM#)): End Sub: End Module"#), vec!["30"]); }
#[test] fn datediff_seconds() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateDiff("s", #12:00:00 PM#, #12:00:45 PM#)): End Sub: End Module"#), vec!["45"]); }
#[test] fn datediff_quarters() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateDiff("q", #1/1/2020#, #4/1/2020#)): End Sub: End Module"#), vec!["1"]); }
#[test] fn datediff_weeks() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateDiff("ww", #1/1/2020#, #1/15/2020#)): End Sub: End Module"#), vec!["2"]); }
#[test] fn datediff_weekdays() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateDiff("w", #1/1/2020#, #1/6/2020#)): End Sub: End Module"#), vec!["5"]); }
#[test] fn datediff_dayofyear() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateDiff("y", #1/1/2020#, #1/11/2020#)): End Sub: End Module"#), vec!["10"]); }

// DatePart tests
#[test] fn datepart_day() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DatePart("d", #1/15/2020#)): End Sub: End Module"#), vec!["15"]); }
#[test] fn datepart_month() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DatePart("m", #5/15/2020#)): End Sub: End Module"#), vec!["5"]); }
#[test] fn datepart_year() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DatePart("yyyy", #5/15/2021#)): End Sub: End Module"#), vec!["2021"]); }
#[test] fn datepart_hour() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DatePart("h", #5:30:00 PM#)): End Sub: End Module"#), vec!["17"]); }
#[test] fn datepart_minute() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DatePart("n", #5:45:00 PM#)): End Sub: End Module"#), vec!["45"]); }
#[test] fn datepart_second() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DatePart("s", #5:45:30 PM#)): End Sub: End Module"#), vec!["30"]); }
#[test] fn datepart_quarter() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DatePart("q", #7/15/2020#)): End Sub: End Module"#), vec!["3"]); }
#[test] fn datepart_week() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DatePart("ww", #1/15/2020#)): End Sub: End Module"#), vec!["3"]); }

// Date creation
#[test] fn date_serial() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateSerial(2020, 5, 10).Month): End Sub: End Module"#), vec!["5"]); }
#[test] fn time_serial() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(TimeSerial(14, 30, 0).Hour): End Sub: End Module"#), vec!["14"]); }
#[test] fn date_value() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(DateValue("2020-05-10").Year): End Sub: End Module"#), vec!["2020"]); }
#[test] fn time_value() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(TimeValue("14:30:00").Minute): End Sub: End Module"#), vec!["30"]); }
#[test] fn date_now_property() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Now.Year > 2000): End Sub: End Module"#), vec!["True"]); }
#[test] fn date_today_property() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Today.Hour): End Sub: End Module"#), vec!["0"]); }
#[test] fn date_timeofday_property() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(TimeOfDay.Year): End Sub: End Module"#), vec!["1"]); }

// Math functions
#[test] fn math_abs_neg() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Abs(-10.5)): End Sub: End Module"#), vec!["10.5"]); }
#[test] fn math_abs_pos() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Abs(10.5)): End Sub: End Module"#), vec!["10.5"]); }
#[test] fn math_sgn_neg() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Sign(-10.5)): End Sub: End Module"#), vec!["-1"]); }
#[test] fn math_sgn_pos() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Sign(10.5)): End Sub: End Module"#), vec!["1"]); }
#[test] fn math_sgn_zero() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Sign(0)): End Sub: End Module"#), vec!["0"]); }
#[test] fn math_fix_pos() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Fix(10.8)): End Sub: End Module"#), vec!["10"]); }
#[test] fn math_fix_neg() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Fix(-10.8)): End Sub: End Module"#), vec!["-10"]); }
#[test] fn math_int_pos() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Int(10.8)): End Sub: End Module"#), vec!["10"]); }
#[test] fn math_int_neg() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Int(-10.8)): End Sub: End Module"#), vec!["-11"]); }
#[test] fn math_round_even() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Round(2.5)): End Sub: End Module"#), vec!["2"]); }
#[test] fn math_round_odd() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Round(3.5)): End Sub: End Module"#), vec!["4"]); }
#[test] fn math_round_digits() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Round(3.14159, 2)): End Sub: End Module"#), vec!["3.14"]); }
#[test] fn math_sqrt() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Sqrt(16)): End Sub: End Module"#), vec!["4"]); }
#[test] fn math_pow() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Pow(2, 3)): End Sub: End Module"#), vec!["8"]); }
#[test] fn math_exp() { assert_eq!(run_vb(r#"Module M: Sub Main(): Console.WriteLine(Math.Round(Math.Exp(1), 2)): End Sub: End Module"#), vec!["2.72"]); }
