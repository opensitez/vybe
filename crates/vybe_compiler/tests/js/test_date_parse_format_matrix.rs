crate::js_cases! {
    invalid_date_tostring_is_invalid_date => {
        r#"
console.log(new Date("bad").toString());
"#,
        ["Invalid Date"]
    };

    invalid_date_toutcstring_is_invalid_date => {
        r#"
console.log(new Date("bad").toUTCString());
"#,
        ["Invalid Date"]
    };

    invalid_date_todatestring_is_invalid_date => {
        r#"
console.log(new Date("bad").toDateString());
"#,
        ["Invalid Date"]
    };

    invalid_date_getfullyear_is_nan => {
        r#"
const d = new Date("bad");
console.log(Number.isNaN(d.getFullYear()));
"#,
        ["true"]
    };

    invalid_date_getutcfullyear_is_nan => {
        r#"
const d = new Date("bad");
console.log(Number.isNaN(d.getUTCFullYear()));
"#,
        ["true"]
    };

    date_parse_rfc_1123_roundtrips_utc_fields => {
        r#"
const d = new Date(Date.parse("Tue, 02 Jan 2024 03:04:05 GMT"));
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
console.log(d.getUTCHours());
"#,
        ["2024", "0", "2", "3"]
    };

    date_parse_positive_offset_converts_to_utc => {
        r#"
const d = new Date(Date.parse("2024-01-02T03:04:05+02:30"));
console.log(d.toISOString());
"#,
        ["2024-01-02T00:34:05.000Z"]
    };

    date_parse_negative_offset_converts_to_utc => {
        r#"
const d = new Date(Date.parse("2024-01-02T03:04:05-05:30"));
console.log(d.toISOString());
"#,
        ["2024-01-02T08:34:05.000Z"]
    };

    date_parse_iso_without_millis_normalizes_to_zero_millis => {
        r#"
const d = new Date(Date.parse("2024-01-02T03:04:05Z"));
console.log(d.getUTCMilliseconds());
"#,
        ["0"]
    };

    date_parse_time_only_string_is_nan => {
        r#"
console.log(Number.isNaN(Date.parse("03:04:05")));
"#,
        ["true"]
    };

    date_toisostring_has_fixed_length_for_common_date => {
        r#"
console.log(new Date(Date.UTC(2024, 0, 2, 3, 4, 5, 6)).toISOString().length);
"#,
        ["24"]
    };

    date_toutcstring_exact_known_value => {
        r#"
console.log(new Date(Date.UTC(2024, 0, 2, 3, 4, 5)).toUTCString());
"#,
        ["Tue, 02 Jan 2024 03:04:05 GMT"]
    };

    date_tojson_exact_known_value => {
        r#"
console.log(new Date(Date.UTC(2024, 0, 2, 3, 4, 5)).toJSON());
"#,
        ["2024-01-02T03:04:05.000Z"]
    };

    date_utc_day_zero_in_constructor_uses_previous_month => {
        r#"
const d = new Date(Date.UTC(2024, 2, 0));
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#,
        ["1", "29"]
    };

    date_local_day_zero_in_constructor_uses_previous_month => {
        r#"
const d = new Date(2024, 2, 0);
console.log(d.getMonth());
console.log(d.getDate());
"#,
        ["1", "29"]
    };

    date_setfullyear_with_month_and_day_updates_all_three => {
        r#"
const d = new Date(2024, 0, 1);
d.setFullYear(2025, 5, 15);
console.log(d.getFullYear());
console.log(d.getMonth());
console.log(d.getDate());
"#,
        ["2025", "5", "15"]
    };

    date_setutcfullyear_with_month_and_day_updates_all_three => {
        r#"
const d = new Date(Date.UTC(2024, 0, 1));
d.setUTCFullYear(2025, 5, 15);
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#,
        ["2025", "5", "15"]
    };

    date_setmonth_on_end_of_month_rolls_forward => {
        r#"
const d = new Date(2024, 0, 31);
d.setMonth(1);
console.log(d.getMonth());
console.log(d.getDate());
"#,
        ["2", "2"]
    };

    date_setutcmonth_on_end_of_month_rolls_forward => {
        r#"
const d = new Date(Date.UTC(2024, 0, 31));
d.setUTCMonth(1);
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#,
        ["2", "2"]
    };

    date_parse_invalid_month_is_nan => {
        r#"
console.log(Number.isNaN(Date.parse("2024-13-01T00:00:00Z")));
"#,
        ["true"]
    };

    date_parse_invalid_day_is_nan => {
        r#"
console.log(Number.isNaN(Date.parse("2024-01-32T00:00:00Z")));
"#,
        ["true"]
    };

    date_parse_fractional_millis_keeps_first_three_digits => {
        r#"
const d = new Date(Date.parse("2024-01-02T03:04:05.6789Z"));
console.log(d.getUTCMilliseconds());
"#,
        ["678"]
    };

    date_parse_epoch_string_roundtrips_to_epoch => {
        r#"
console.log(new Date(Date.parse("1970-01-01T00:00:00Z")).getTime());
"#,
        ["0"]
    };
}