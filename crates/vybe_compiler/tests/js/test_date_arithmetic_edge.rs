//! Date construction, UTC/local getters, and arithmetic edge cases.

crate::js_cases! {
    date_constructor_no_args_is_now => {
        r#"console.log(new Date()>new Date(0));"#,
        ["true"]
    };

    date_constructor_milliseconds => {
        r#"console.log(new Date(0).getTime());"#,
        ["0"]
    };

    date_constructor_iso_string => {
        r#"console.log(new Date("1970-01-01T00:00:00.000Z").getUTCFullYear());"#,
        ["1970"]
    };

    date_constructor_year_month_day => {
        r#"console.log(new Date(2024,0,15).getDate());"#,
        ["15"]
    };

    date_get_full_year => {
        r#"console.log(new Date(2024,5,1).getFullYear());"#,
        ["2024"]
    };

    date_get_month_zero_indexed => {
        r#"console.log(new Date(2024,0,1).getMonth());"#,
        ["0"]
    };

    date_get_utc_full_year => {
        r#"console.log(new Date("2024-06-01T00:00:00Z").getUTCFullYear());"#,
        ["2024"]
    };

    date_get_utc_month => {
        r#"console.log(new Date("2024-06-01T00:00:00Z").getUTCMonth());"#,
        ["5"]
    };

    date_get_time_epoch_ms => {
        r#"console.log(new Date("1970-01-01T00:00:00.000Z").getTime());"#,
        ["0"]
    };

    date_set_full_year_mutates => {
        r#"const d=new Date(2020,0,1); d.setFullYear(2025); console.log(d.getFullYear());"#,
        ["2025"]
    };

    date_set_month_mutates => {
        r#"const d=new Date(2024,0,1); d.setMonth(11); console.log(d.getMonth());"#,
        ["11"]
    };

    date_set_date_mutates => {
        r#"const d=new Date(2024,0,10); d.setDate(20); console.log(d.getDate());"#,
        ["20"]
    };

    date_toisostring_utc_format => {
        r#"console.log(new Date(0).toISOString().startsWith("1970"));"#,
        ["true"]
    };

    date_toutcstring_includes_gmt => {
        r#"console.log(new Date(0).toUTCString().includes("1970"));"#,
        ["true"]
    };

    date_tojson_is_quoted_iso => {
        r#"console.log(JSON.parse(JSON.stringify(new Date(0))).startsWith("1970"));"#,
        ["true"]
    };

    date_get_day_of_week => {
        r#"console.log(new Date("2024-01-07T00:00:00Z").getUTCDay());"#,
        ["0"]
    };

    date_valueof_equals_gettime => {
        r#"const d=new Date(1000); console.log(d.valueOf()===d.getTime());"#,
        ["true"]
    };

    date_addition_coerces_to_number => {
        r#"console.log(new Date(0)+1>0);"#,
        ["true"]
    };

    date_invalid_string_yields_invalid_date => {
        r#"console.log(Number.isNaN(new Date("not-a-date").getTime()));"#,
        ["true"]
    };

    date_compare_less_than => {
        r#"console.log(new Date(0)<new Date(1));"#,
        ["true"]
    };

    date_compare_equality_same_ms => {
        r#"console.log(new Date(5).getTime()===new Date(5).getTime());"#,
        ["true"]
    };

    date_get_hours_local => {
        r#"const d=new Date(2024,0,1,13,0,0); console.log(d.getHours());"#,
        ["13"]
    };

    date_get_minutes => {
        r#"const d=new Date(2024,0,1,0,45,0); console.log(d.getMinutes());"#,
        ["45"]
    };

    date_get_seconds => {
        r#"const d=new Date(2024,0,1,0,0,30); console.log(d.getSeconds());"#,
        ["30"]
    };

    date_get_milliseconds => {
        r#"const d=new Date(2024,0,1,0,0,0,250); console.log(d.getMilliseconds());"#,
        ["250"]
    };

    date_set_time_updates_all_fields => {
        r#"const d=new Date(0); d.setTime(86400000); console.log(d.getUTCDate());"#,
        ["2"]
    };

    date_parse_rfc_string => {
        r#"console.log(Date.parse("Jan 1, 2024")>0);"#,
        ["true"]
    };

    date_utc_constructor => {
        r#"console.log(Date.UTC(2024,0,1));"#,
        ["1704067200000"]
    };

    date_now_monotonic_increase => {
        r#"const a=Date.now(); const b=Date.now(); console.log(b>=a);"#,
        ["true"]
    };

    date_timezone_offset_is_minutes => {
        r#"console.log(typeof new Date().getTimezoneOffset());"#,
        ["number"]
    };

    date_instanceof_date => {
        r#"console.log(new Date() instanceof Date);"#,
        ["true"]
    };

    date_prototype_to_string_not_empty => {
        r#"console.log(new Date(0).toString().length>0);"#,
        ["true"]
    };

    date_set_utc_full_year => {
        r#"const d=new Date(0); d.setUTCFullYear(2021); console.log(d.getUTCFullYear());"#,
        ["2021"]
    };

    date_set_utc_month => {
        r#"const d=new Date(0); d.setUTCMonth(6); console.log(d.getUTCMonth());"#,
        ["6"]
    };

    date_set_utc_date => {
        r#"const d=new Date(0); d.setUTCDate(10); console.log(d.getUTCDate());"#,
        ["10"]
    };

    date_leap_year_february => {
        r#"console.log(new Date(2024,1,29).getDate());"#,
        ["29"]
    };

    date_year_rollover_december_to_january => {
        r#"const d=new Date(2023,11,31); d.setDate(d.getDate()+1); console.log(d.getMonth());"#,
        ["0"]
    };

    date_diff_in_ms_subtraction => {
        r#"console.log(new Date(1000)-new Date(0));"#,
        ["1000"]
    };

    date_from_timestamp_roundtrip => {
        r#"const t=1234567890123; console.log(new Date(t).getTime());"#,
        ["1234567890123"]
    };

    date_to_date_string_contains_weekday => {
        r#"console.log(new Date("2024-01-01T00:00:00Z").toDateString().length>6);"#,
        ["true"]
    };

    date_to_time_string_contains_colon => {
        r#"console.log(new Date(2024,0,1,12,30,0).toTimeString().includes(":"));"#,
        ["true"]
    };

    date_set_hours_overflow_rolls_day => {
        r#"const d=new Date(2024,0,1,0,0,0); d.setHours(25); console.log(d.getDate());"#,
        ["2"]
    };

    date_get_timezone_offset_sign => {
        r#"console.log(typeof new Date("2024-01-01").getTimezoneOffset());"#,
        ["number"]
    };

    date_constructor_date_only_string_utc_midnight => {
        r#"console.log(new Date("2024-01-01").toISOString().includes("2024-01-01"));"#,
        ["true"]
    };
}
