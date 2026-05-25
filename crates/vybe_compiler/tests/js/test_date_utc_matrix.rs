crate::js_cases! {
    date_utc_constructor_exposes_full_year_month_day => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2));
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#,
        ["2024", "0", "2"]
    };

    date_utc_constructor_exposes_time_fields => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 3, 4, 5, 6));
console.log(d.getUTCHours());
console.log(d.getUTCMinutes());
console.log(d.getUTCSeconds());
console.log(d.getUTCMilliseconds());
"#,
        ["3", "4", "5", "6"]
    };

    date_getutcday_for_known_tuesday => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2));
console.log(d.getUTCDay());
"#,
        ["2"]
    };

    date_valueof_matches_gettime => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 3, 4, 5, 6));
console.log(d.valueOf() === d.getTime());
"#,
        ["true"]
    };

    date_tojson_matches_toisostring_for_valid_date => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 3, 4, 5, 6));
console.log(d.toJSON() === d.toISOString());
"#,
        ["true"]
    };

    invalid_date_gettime_is_nan => {
        r#"
const d = new Date("not a date");
console.log(Number.isNaN(d.getTime()));
"#,
        ["true"]
    };

    invalid_date_tojson_is_null => {
        r#"
const d = new Date("not a date");
console.log(d.toJSON() === null);
"#,
        ["true"]
    };

    date_setutcfullyear_updates_year => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2));
d.setUTCFullYear(2030);
console.log(d.getUTCFullYear());
"#,
        ["2030"]
    };

    date_setutcmonth_updates_month => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2));
d.setUTCMonth(6);
console.log(d.getUTCMonth());
"#,
        ["6"]
    };

    date_setutcdate_updates_day => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2));
d.setUTCDate(15);
console.log(d.getUTCDate());
"#,
        ["15"]
    };

    date_setutchours_updates_hour => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 1, 2, 3));
d.setUTCHours(22);
console.log(d.getUTCHours());
"#,
        ["22"]
    };

    date_setutcminutes_updates_minute => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 1, 2, 3));
d.setUTCMinutes(44);
console.log(d.getUTCMinutes());
"#,
        ["44"]
    };

    date_setutcseconds_updates_second => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 1, 2, 3));
d.setUTCSeconds(55);
console.log(d.getUTCSeconds());
"#,
        ["55"]
    };

    date_setutcmilliseconds_updates_millis => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 1, 2, 3, 4));
d.setUTCMilliseconds(99);
console.log(d.getUTCMilliseconds());
"#,
        ["99"]
    };

    date_setutcmonth_overflow_rolls_year => {
        r#"
const d = new Date(Date.UTC(2024, 10, 2));
d.setUTCMonth(12);
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
"#,
        ["2025", "0"]
    };

    date_setutcdate_overflow_rolls_month => {
        r#"
const d = new Date(Date.UTC(2024, 0, 31));
d.setUTCDate(32);
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#,
        ["1", "1"]
    };

    date_setutcdate_zero_moves_to_previous_month => {
        r#"
const d = new Date(Date.UTC(2024, 2, 1));
d.setUTCDate(0);
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#,
        ["1", "29"]
    };

    date_setutchours_overflow_rolls_day => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 23, 0, 0));
d.setUTCHours(24);
console.log(d.getUTCDate());
console.log(d.getUTCHours());
"#,
        ["3", "0"]
    };

    date_setutcminutes_overflow_rolls_hour => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 1, 59, 0));
d.setUTCMinutes(60);
console.log(d.getUTCHours());
console.log(d.getUTCMinutes());
"#,
        ["2", "0"]
    };

    date_setutcseconds_overflow_rolls_minute => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 1, 2, 59));
d.setUTCSeconds(60);
console.log(d.getUTCMinutes());
console.log(d.getUTCSeconds());
"#,
        ["3", "0"]
    };

    date_setutcmilliseconds_overflow_rolls_second => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 1, 2, 3, 999));
d.setUTCMilliseconds(1000);
console.log(d.getUTCSeconds());
console.log(d.getUTCMilliseconds());
"#,
        ["4", "0"]
    };

    date_utc_month_overflow_in_constructor_rolls_year => {
        r#"
const d = new Date(Date.UTC(2024, 12, 1));
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
"#,
        ["2025", "0"]
    };

    date_utc_negative_month_in_constructor_rolls_previous_year => {
        r#"
const d = new Date(Date.UTC(2024, -1, 1));
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
"#,
        ["2023", "11"]
    };

    date_parse_iso_with_milliseconds => {
        r#"
const ts = Date.parse("2024-01-02T03:04:05.006Z");
const d = new Date(ts);
console.log(d.getUTCSeconds());
console.log(d.getUTCMilliseconds());
"#,
        ["5", "6"]
    };

    date_parse_date_only_uses_midnight_utc => {
        r#"
const d = new Date(Date.parse("2024-01-02"));
console.log(d.toISOString());
"#,
        ["2024-01-02T00:00:00.000Z"]
    };

    date_toisostring_for_epoch_zero => {
        r#"
console.log(new Date(0).toISOString());
"#,
        ["1970-01-01T00:00:00.000Z"]
    };

    date_toisostring_for_negative_one_millisecond => {
        r#"
console.log(new Date(-1).toISOString());
"#,
        ["1969-12-31T23:59:59.999Z"]
    };

    date_tojson_for_epoch_zero => {
        r#"
console.log(new Date(0).toJSON());
"#,
        ["1970-01-01T00:00:00.000Z"]
    };

    date_settime_updates_timestamp => {
        r#"
const d = new Date(0);
d.setTime(1000);
console.log(d.getTime());
console.log(d.getUTCSeconds());
"#,
        ["1000", "1"]
    };

    date_gettimezoneoffset_returns_number => {
        r#"
const d = new Date();
console.log(typeof d.getTimezoneOffset());
"#,
        ["number"]
    };

    date_toutcstring_contains_gmt => {
        r#"
const d = new Date(Date.UTC(2024, 0, 2, 3, 4, 5));
console.log(d.toUTCString().includes("GMT"));
"#,
        ["true"]
    };

    date_totimestring_contains_colons => {
        r#"
const d = new Date(0);
const s = d.toTimeString();
console.log(s.includes(":"));
"#,
        ["true"]
    };

    date_todatestring_is_non_empty => {
        r#"
const d = new Date(0);
console.log(d.toDateString().length > 0);
"#,
        ["true"]
    };

    date_utc_leap_day_is_preserved => {
        r#"
const d = new Date(Date.UTC(2024, 1, 29));
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#,
        ["1", "29"]
    };

    date_local_setdate_overflow_rolls_month => {
        r#"
const d = new Date(2024, 0, 31);
d.setDate(32);
console.log(d.getMonth());
console.log(d.getDate());
"#,
        ["1", "1"]
    };

    date_local_setmonth_overflow_rolls_year => {
        r#"
const d = new Date(2024, 11, 1);
d.setMonth(12);
console.log(d.getFullYear());
console.log(d.getMonth());
"#,
        ["2025", "0"]
    };

    date_local_sethours_overflow_rolls_date => {
        r#"
const d = new Date(2024, 0, 1, 23, 0, 0);
d.setHours(24);
console.log(d.getDate());
console.log(d.getHours());
"#,
        ["2", "0"]
    };

    date_local_setminutes_overflow_rolls_hour => {
        r#"
const d = new Date(2024, 0, 1, 1, 59, 0);
d.setMinutes(60);
console.log(d.getHours());
console.log(d.getMinutes());
"#,
        ["2", "0"]
    };

    date_local_setseconds_overflow_rolls_minute => {
        r#"
const d = new Date(2024, 0, 1, 1, 2, 59);
d.setSeconds(60);
console.log(d.getMinutes());
console.log(d.getSeconds());
"#,
        ["3", "0"]
    };

    date_local_setmilliseconds_overflow_rolls_second => {
        r#"
const d = new Date(2024, 0, 1, 1, 2, 3, 999);
d.setMilliseconds(1000);
console.log(d.getSeconds());
console.log(d.getMilliseconds());
"#,
        ["4", "0"]
    };

    date_constructor_from_iso_z_preserves_utc_components => {
        r#"
const d = new Date("2024-01-02T03:04:05Z");
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
console.log(d.getUTCHours());
"#,
        ["2024", "0", "2", "3"]
    };

    date_constructor_numeric_timestamp_preserves_milliseconds => {
        r#"
const d = new Date(123456789);
console.log(d.getTime());
"#,
        ["123456789"]
    };

    date_utc_constructor_accepts_only_year_month => {
        r#"
const d = new Date(Date.UTC(2024, 0));
console.log(d.getUTCDate());
console.log(d.getUTCHours());
"#,
        ["1", "0"]
    };
}