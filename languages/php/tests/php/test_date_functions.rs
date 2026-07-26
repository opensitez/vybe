//! `date`, `strtotime`, `DateTime`, `DateInterval`, and timezone helpers — fixed UTC outputs.

crate::php_cases! {
    date_formats_fixed_timestamp_as_year_month_day => {
        r#"<?php
date_default_timezone_set('UTC');
echo date('Y-m-d', 1704067200);
"#,
        ["2024-01-01"]
    };

    date_extracts_hour_minute_from_timestamp => {
        r#"<?php
date_default_timezone_set('UTC');
echo date('H:i', 1718461800);
"#,
        ["14:30"]
    };

    mktime_builds_midnight_march_fifteenth => {
        r#"<?php
date_default_timezone_set('UTC');
echo date('Y-m-d', mktime(0, 0, 0, 3, 15, 2024));
"#,
        ["2024-03-15"]
    };

    strtotime_parses_iso_date_string => {
        r#"<?php
date_default_timezone_set('UTC');
echo date('Y', strtotime('2024-06-15'));
"#,
        ["2024"]
    };

    strtotime_relative_plus_one_week => {
        r#"<?php
date_default_timezone_set('UTC');
$base = strtotime('2024-01-01');
echo date('Y-m-d', strtotime('+1 week', $base));
"#,
        ["2024-01-08"]
    };

    strtotime_relative_plus_one_month => {
        r#"<?php
date_default_timezone_set('UTC');
$base = strtotime('2024-01-15');
echo date('m', strtotime('+1 month', $base));
"#,
        ["02"]
    };

    getdate_returns_expected_mday_and_mon => {
        r#"<?php
date_default_timezone_set('UTC');
$info = getdate(mktime(10, 30, 0, 7, 4, 2024));
echo $info['mon'] . ':' . $info['mday'];
"#,
        ["7:4"]
    };

    checkdate_accepts_leap_day_2024 => {
        r#"<?php
echo checkdate(2, 29, 2024) ? 'leap' : 'no';
"#,
        ["leap"]
    };

    checkdate_rejects_feb_29_non_leap => {
        r#"<?php
echo checkdate(2, 29, 2023) ? 'leap' : 'no';
"#,
        ["no"]
    };

    checkdate_rejects_month_thirteen => {
        r#"<?php
echo checkdate(13, 1, 2024) ? 'ok' : 'bad';
"#,
        ["bad"]
    };

    date_create_parses_christmas => {
        r#"<?php
echo date_format(date_create('2024-12-25'), 'm-d');
"#,
        ["12-25"]
    };

    date_modify_adds_three_months => {
        r#"<?php
$dt = date_create('2024-01-01');
date_modify($dt, '+3 months');
echo date_format($dt, 'Y-m');
"#,
        ["2024-04"]
    };

    datetime_timezone_conversion_keeps_instant => {
        r#"<?php
$source = new DateTime('2024-03-01 12:00:00', new DateTimeZone('UTC'));
$local = $source->setTimezone(new DateTimeZone('Europe/Paris'));
echo $source->format('H:i') . '|' . $local->format('H:i');
"#,
        ["12:00|13:00"]
    };

    date_diff_counts_days_between_dates => {
        r#"<?php
$d1 = date_create('2024-01-01');
$d2 = date_create('2024-01-11');
echo (int)date_diff($d1, $d2)->days;
"#,
        ["10"]
    };

    date_add_interval_two_months => {
        r#"<?php
$dt = date_create('2024-01-01');
date_add($dt, new DateInterval('P2M'));
echo date_format($dt, 'Y-m');
"#,
        ["2024-03"]
    };

    date_sub_interval_fifteen_days => {
        r#"<?php
$dt = date_create('2024-06-15');
date_sub($dt, new DateInterval('P15D'));
echo date_format($dt, 'Y-m-d');
"#,
        ["2024-05-31"]
    };

    date_interval_from_three_weeks_string => {
        r#"<?php
$i = date_interval_create_from_date_string('3 weeks');
echo ($i instanceof DateInterval) ? 'ok' : 'fail';
"#,
        ["ok"]
    };

    date_timestamp_get_reads_unix_time => {
        r#"<?php
date_default_timezone_set('UTC');
$dt = date_create('2024-01-01 00:00:00');
echo date('Y', date_timestamp_get($dt));
"#,
        ["2024"]
    };

    date_timestamp_set_overwrites_datetime => {
        r#"<?php
date_default_timezone_set('UTC');
$dt = date_create('2000-01-01');
date_timestamp_set($dt, mktime(0, 0, 0, 6, 15, 2024));
echo date_format($dt, 'Y-m-d');
"#,
        ["2024-06-15"]
    };

    datetime_parsed_from_timezoneaware_iso_string => {
        r#"<?php
date_default_timezone_set('UTC');
$dt = new DateTime('2024-07-01T10:15:00+02:00');
echo $dt->format('Y-m-d H:i');
echo '|';
echo $dt->getTimezone()->getName();
"#,
        ["2024-07-01 10:15|+02:00"]
    };

    datetime_set_timezone_preserves_instant => {
        r#"<?php
$dt = new DateTime('2024-01-01 12:00:00', new DateTimeZone('Europe/Paris'));
echo $dt->format('H') . '|';
$dt->setTimezone(new DateTimeZone('UTC'));
echo $dt->format('H') . '|' . $dt->getTimezone()->getName();
"#,
        ["12|11|UTC"]
    };

    datetime_createfromformat_parses_dmy => {
        r#"<?php
$dt = DateTime::createFromFormat('d/m/Y', '25/12/2024');
echo $dt->format('Y-m-d');
"#,
        ["2024-12-25"]
    };

    datetime_modify_hours_crossing_midnight => {
        r#"<?php
$dt = new DateTime('2024-06-01 23:30:00');
$dt->modify('+90 minutes');
echo $dt->format('Y-m-d H:i');
"#,
        ["2024-06-02 01:00"]
    };

    datetime_format_day_name_for_known_date => {
        r#"<?php
date_default_timezone_set('UTC');
$dt = new DateTime('2024-06-15');
echo $dt->format('N');
"#,
        ["6"]
    };

    datetime_modify_minus_one_day => {
        r#"<?php
$dt = new DateTime('2024-06-02');
$dt->modify('-1 day');
echo $dt->format('d');
"#,
        ["01"]
    };

    datetime_diff_months_between_jan_and_jun => {
        r#"<?php
$a = new DateTime('2024-01-01');
$b = new DateTime('2024-06-01');
echo (int)date_diff($a, $b)->m;
"#,
        ["5"]
    };

    datetimeimmutable_add_returns_new_instance => {
        r#"<?php
$orig = new DateTimeImmutable('2024-01-01');
$new = $orig->add(new DateInterval('P1D'));
echo $orig->format('d') . ':' . $new->format('d');
"#,
        ["01:02"]
    };

    datetime_set_time_components => {
        r#"<?php
$dt = new DateTime('2024-01-01');
$dt->setTime(9, 5, 0);
echo $dt->format('H:i:s');
"#,
        ["09:05:00"]
    };

    datetime_set_date_components => {
        r#"<?php
$dt = new DateTime('2024-01-01');
$dt->setDate(2025, 12, 31);
echo $dt->format('Y-m-d');
"#,
        ["2025-12-31"]
    };

    gmdate_utc_independent_of_local_tz => {
        r#"<?php
date_default_timezone_set('America/New_York');
echo gmdate('Y-m-d', 1704067200);
"#,
        ["2024-01-01"]
    };

    idate_returns_numeric_month_from_timestamp => {
        r#"<?php
date_default_timezone_set('UTC');
echo idate('m', strtotime('2024-06-15'));
"#,
        ["6"]
    };

    strftime_style_via_date_format_week => {
        r#"<?php
date_default_timezone_set('UTC');
echo date('W', strtotime('2024-01-01'));
"#,
        ["01"]
    };

    timezone_open_europe_paris_name => {
        r#"<?php
echo timezone_open('Europe/Paris')->getName();
"#,
        ["Europe/Paris"]
    };

    datetimezone_get_offset_for_utc => {
        r#"<?php
$tz = new DateTimeZone('UTC');
$dt = new DateTime('2024-06-01', $tz);
echo $tz->getOffset($dt);
"#,
        ["0"]
    };

    date_default_timezone_get_returns_string => {
        r#"<?php
echo is_string(date_default_timezone_get()) ? 'tz' : 'no';
"#,
        ["tz"]
    };

    microtime_true_returns_float_greater_than_epoch => {
        r#"<?php
echo microtime(true) > 1_000_000_000 ? 'float' : 'small';
"#,
        ["float"]
    };

    date_parse_strictly_parses_rfc3339_string => {
        r#"<?php
$dt = date_parse('2024-07-01T15:30:00Z');
echo $dt['year'] . '-' . str_pad((string) $dt['month'], 2, '0', STR_PAD_LEFT) . '-' . str_pad((string) $dt['day'], 2, '0', STR_PAD_LEFT);
"#,
        ["2024-07-01"]
    };

    date_parse_invalid_input_reports_error_count => {
        r#"<?php
$parsed = date_parse('not-a-date');
echo is_array($parsed) && ($parsed['error_count'] ?? 0) > 0 ? 'err' : 'ok';
"#,
        ["err"]
    };

    timezone_name_from_abbr_lookup => {
        r#"<?php
echo timezone_name_from_abbr('UTC') === 'UTC' ? 'UTC' : 'wrong';
"#,
        ["UTC"]
    };

    date_timestamp_get_of_local_datetime => {
        r#"<?php
date_default_timezone_set('UTC');
$dt = new DateTime('2024-01-01 00:00:00');
echo (int)($dt->getTimestamp() === 1704067200 ? 1 : 0);
"#,
        ["1"]
    };

    timezone_identifier_is_listed_or_ignored => {
        r#"<?php
$ids = DateTimeZone::listIdentifiers(DateTimeZone::AFRICA);
echo in_array('Africa/Casablanca', $ids) ? 'has' : 'missing';
"#,
        ["has"]
    };

    date_weekday_name_for_iso_monday => {
        r#"<?php
date_default_timezone_set('UTC');
echo date('l', strtotime('2024-02-05'));
"#,
        ["Monday"]
    };

    date_idate_weekday_number => {
        r#"<?php
date_default_timezone_set('UTC');
echo idate('w', strtotime('2024-02-03'));
"#,
        ["6"]
    };

    date_get_last_errors_empty_for_valid => {
        r#"<?php
date_parse('2024-12-01');
echo is_array(date_get_last_errors()) ? 'ok' : 'bad';
echo '|';
echo date_get_last_errors()['warning_count'];
echo '|';
echo date_get_last_errors()['error_count'];
"#,
        ["ok|0|0"]
    };

    date_create_from_format_with_timezone_object => {
        r#"<?php
date_default_timezone_set('UTC');
$tz = new DateTimeZone('America/Los_Angeles');
$dt = date_create_from_format('Y-m-d H:i', '2024-11-05 01:30', $tz);
echo $dt->format('e');
echo '|';
echo $dt->format('H:i');
"#,
        ["America/Los_Angeles|01:30"]
    };

    strtotime_with_trailing_timezone_abbr => {
        r#"<?php
date_default_timezone_set('UTC');
$ts = strtotime('2024-06-01 12:00 UTC');
echo date('H', $ts);
"#,
        ["12"]
    };

    gmmktime_vs_mktime_unix_midday => {
        r#"<?php
date_default_timezone_set('America/New_York');
$g = gmmktime(12, 0, 0, 1, 1, 2024);
$l = mktime(12, 0, 0, 1, 1, 2024);
echo date('H', $g);
echo '|';
echo date('H', $l);
"#,
        ["12|17"]
    };

    timezone_transitions_for_single_range => {
        r#"<?php
$tz = new DateTimeZone('America/New_York');
$transitions = $tz->getTransitions(1704067200, 1706745600);
echo is_array($transitions) ? 'array' : 'bad';
echo '|';
echo count($transitions) > 1 ? 'many' : 'few';
"#,
        ["array|many"]
    };
}
