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

    datetime_createfromformat_parses_dmy => {
        r#"<?php
$dt = DateTime::createFromFormat('d/m/Y', '25/12/2024');
echo $dt->format('Y-m-d');
"#,
        ["2024-12-25"]
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
}
