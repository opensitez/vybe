//! Invalid date/time parsing — `createFromFormat`, constructors, `checkdate`, `strtotime`.

crate::php_cases! {
    datetime_create_from_format_garbage_returns_false => {
        r#"<?php
$dt = DateTime::createFromFormat('Y-m-d', 'not-a-date');
echo $dt === false ? 'false' : 'ok';
"#,
        ["false"]
    };

    datetime_get_last_errors_after_bad_create => {
        r#"<?php
DateTime::createFromFormat('Y-m-d', 'bad');
$e = DateTime::getLastErrors();
echo ($e['error_count'] ?? 0) > 0 ? 'errs' : 'clean';
"#,
        ["errs"]
    };

    datetime_constructor_invalid_string_throws => {
        r#"<?php
try { new DateTime('totally-invalid'); echo 'ok'; }
catch (Exception $e) { echo 'ctor'; }
"#,
        ["ctor"]
    };

    datetime_immutable_invalid_string_throws => {
        r#"<?php
try { new DateTimeImmutable('nope'); echo 'ok'; }
catch (Exception $e) { echo 'immut'; }
"#,
        ["immut"]
    };

    datetime_create_from_format_impossible_calendar_date => {
        r#"<?php
$dt = DateTime::createFromFormat('Y-m-d', '2023-02-30');
echo $dt === false ? 'impossible' : 'ok';
"#,
        ["impossible"]
    };

    datetime_create_from_format_strict_day_month_swap => {
        r#"<?php
$dt = DateTime::createFromFormat('d/m/Y', '31/02/2024');
echo $dt === false ? 'swap-fail' : 'ok';
"#,
        ["swap-fail"]
    };

    datetime_create_from_format_empty_string_returns_false => {
        r#"<?php
$dt = DateTime::createFromFormat('Y-m-d', '');
echo $dt === false ? 'empty' : 'ok';
"#,
        ["empty"]
    };

    datetime_create_from_format_trailing_junk_sets_warning => {
        r#"<?php
$dt = DateTime::createFromFormat('Y-m-d', '2024-01-01 extra');
$errs = DateTime::getLastErrors();
echo ($dt !== false && ($errs['warning_count'] ?? 0) > 0) ? 'warn' : 'none';
"#,
        ["warn"]
    };

    datetime_get_last_errors_clean_after_valid_create => {
        r#"<?php
DateTime::createFromFormat('Y-m-d', '2024-06-01');
$e = DateTime::getLastErrors();
echo (($e['error_count'] ?? 0) + ($e['warning_count'] ?? 0)) === 0 ? 'clean' : 'dirty';
"#,
        ["clean"]
    };

    checkdate_rejects_february_thirtieth => {
        r#"<?php
echo checkdate(2, 30, 2024) ? 'valid' : 'invalid';
"#,
        ["invalid"]
    };

    checkdate_rejects_month_thirteen => {
        r#"<?php
echo checkdate(13, 1, 2024) ? 'valid' : 'invalid';
"#,
        ["invalid"]
    };

    checkdate_rejects_day_zero => {
        r#"<?php
echo checkdate(1, 0, 2024) ? 'valid' : 'invalid';
"#,
        ["invalid"]
    };

    checkdate_accepts_leap_day_in_leap_year => {
        r#"<?php
echo checkdate(2, 29, 2024) ? 'leap' : 'no';
"#,
        ["leap"]
    };

    checkdate_rejects_feb_twenty_ninth_non_leap => {
        r#"<?php
echo checkdate(2, 29, 2023) ? 'leap' : 'no';
"#,
        ["no"]
    };

    strtotime_invalid_phrase_returns_false => {
        r#"<?php
$t = strtotime('not a real time phrase');
echo $t === false ? 'false' : 'set';
"#,
        ["false"]
    };

    strtotime_empty_string_returns_false => {
        r#"<?php
$t = strtotime('');
echo $t === false ? 'false' : 'set';
"#,
        ["false"]
    };

    datetime_modify_with_invalid_modifier_string_returns_false => {
        r#"<?php
$d = new DateTime('2024-01-01');
$ok = $d->modify('not a valid modifier');
echo $ok === false ? 'bad-mod' : 'ok';
"#,
        ["bad-mod"]
    };

    datetime_diff_inverted_flag_when_earlier => {
        r#"<?php
$a = new DateTime('2024-06-10');
$b = new DateTime('2024-06-01');
echo $a->diff($b)->invert ? 'inv' : 'fwd';
"#,
        ["inv"]
    };

    datetime_create_from_format_wrong_separator_returns_false => {
        r#"<?php
$dt = DateTime::createFromFormat('Y-m-d', '2024/06/01');
echo $dt === false ? 'sep' : 'ok';
"#,
        ["sep"]
    };

    datetime_immutable_create_from_format_bad_input => {
        r#"<?php
$dt = DateTimeImmutable::createFromFormat('H:i:s', '25:99:99');
echo $dt === false ? 'bad-time' : 'ok';
"#,
        ["bad-time"]
    };
}
