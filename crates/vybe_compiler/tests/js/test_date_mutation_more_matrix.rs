crate::js_cases! {
    date_sethours_with_minutes_and_seconds_updates_all => {
        r#"
const d = new Date(2024, 0, 1, 1, 2, 3);
d.setHours(4, 5, 6, 7);
console.log(d.getHours());
console.log(d.getMinutes());
console.log(d.getSeconds());
console.log(d.getMilliseconds());
"#,
        ["4", "5", "6", "7"]
    };
    date_setminutes_with_seconds_and_millis_updates_all => {
        r#"
const d = new Date(2024, 0, 1, 1, 2, 3, 4);
d.setMinutes(10, 11, 12);
console.log(d.getMinutes());
console.log(d.getSeconds());
console.log(d.getMilliseconds());
"#,
        ["10", "11", "12"]
    };
    date_setseconds_with_millis_updates_both => {
        r#"
const d = new Date(2024, 0, 1, 1, 2, 3, 4);
d.setSeconds(20, 21);
console.log(d.getSeconds());
console.log(d.getMilliseconds());
"#,
        ["20", "21"]
    };
    date_setutchours_with_minutes_and_seconds_updates_all => {
        r#"
const d = new Date(Date.UTC(2024, 0, 1, 1, 2, 3, 4));
d.setUTCHours(4, 5, 6, 7);
console.log(d.getUTCHours());
console.log(d.getUTCMinutes());
console.log(d.getUTCSeconds());
console.log(d.getUTCMilliseconds());
"#,
        ["4", "5", "6", "7"]
    };
    date_setutcminutes_with_seconds_and_millis_updates_all => {
        r#"
const d = new Date(Date.UTC(2024, 0, 1, 1, 2, 3, 4));
d.setUTCMinutes(10, 11, 12);
console.log(d.getUTCMinutes());
console.log(d.getUTCSeconds());
console.log(d.getUTCMilliseconds());
"#,
        ["10", "11", "12"]
    };
    date_setutcseconds_with_millis_updates_both => {
        r#"
const d = new Date(Date.UTC(2024, 0, 1, 1, 2, 3, 4));
d.setUTCSeconds(20, 21);
console.log(d.getUTCSeconds());
console.log(d.getUTCMilliseconds());
"#,
        ["20", "21"]
    };
    date_setfull_year_returns_timestamp_number => {
        r#"
const d = new Date(2024, 0, 1);
console.log(typeof d.setFullYear(2025));
"#,
        ["number"]
    };
    date_setutcfullyear_returns_timestamp_number => {
        r#"
const d = new Date(Date.UTC(2024, 0, 1));
console.log(typeof d.setUTCFullYear(2025));
"#,
        ["number"]
    };
    date_setmonth_returns_timestamp_number => {
        r#"
const d = new Date(2024, 0, 1);
console.log(typeof d.setMonth(5));
"#,
        ["number"]
    };
    date_setutcmonth_returns_timestamp_number => {
        r#"
const d = new Date(Date.UTC(2024, 0, 1));
console.log(typeof d.setUTCMonth(5));
"#,
        ["number"]
    };
    date_getyear_is_full_year_minus_1900 => {
        r#"
console.log(new Date(2024, 0, 1).getYear());
"#,
        ["124"]
    };
    date_now_returns_number => {
        r#"
console.log(typeof Date.now());
"#,
        ["number"]
    };
    date_parse_returns_number_for_valid_iso => {
        r#"
console.log(typeof Date.parse("2024-01-01T00:00:00Z"));
"#,
        ["number"]
    };
    date_constructor_without_arguments_produces_valid_date => {
        r#"
const d = new Date();
console.log(!Number.isNaN(d.getTime()));
"#,
        ["true"]
    };
    date_constructor_copy_preserves_timestamp => {
        r#"
const a = new Date(1234);
const b = new Date(a);
console.log(a.getTime() === b.getTime());
"#,
        ["true"]
    };
    date_setdate_returns_timestamp_number => {
        r#"
const d = new Date(2024, 0, 1);
console.log(typeof d.setDate(2));
"#,
        ["number"]
    };
    date_setutcdate_returns_timestamp_number => {
        r#"
const d = new Date(Date.UTC(2024, 0, 1));
console.log(typeof d.setUTCDate(2));
"#,
        ["number"]
    };
    date_settime_returns_assigned_timestamp => {
        r#"
const d = new Date(0);
console.log(d.setTime(1234));
"#,
        ["1234"]
    };
    date_valueof_on_invalid_date_is_nan => {
        r#"
console.log(Number.isNaN(new Date("bad").valueOf()));
"#,
        ["true"]
    };
    date_getday_for_utc_epoch_day_matches_known_weekday => {
        r#"
console.log(new Date(0).getUTCDay());
"#,
        ["4"]
    };
    date_month_getter_zero_indexes_january => {
        r#"
console.log(new Date(2024, 0, 1).getMonth());
"#,
        ["0"]
    };
    date_utc_month_getter_zero_indexes_january => {
        r#"
console.log(new Date(Date.UTC(2024, 0, 1)).getUTCMonth());
"#,
        ["0"]
    };
    date_setmilliseconds_returns_timestamp_number => {
        r#"
const d = new Date(0);
console.log(typeof d.setMilliseconds(5));
"#,
        ["number"]
    };
    date_setutcmilliseconds_returns_timestamp_number => {
        r#"
const d = new Date(0);
console.log(typeof d.setUTCMilliseconds(5));
"#,
        ["number"]
    };
    date_constructor_numeric_string_parses_as_date_string_not_timestamp_number => {
        r#"
const a = new Date("1234");
console.log(!Number.isNaN(a.getTime()));
"#,
        ["true"]
    };
    date_tostring_on_epoch_is_non_empty => {
        r#"
console.log(new Date(0).toString().length > 0);
"#,
        ["true"]
    };
    date_toutcstring_on_epoch_exact_value => {
        r#"
console.log(new Date(0).toUTCString());
"#,
        ["Thu, 01 Jan 1970 00:00:00 GMT"]
    };
}