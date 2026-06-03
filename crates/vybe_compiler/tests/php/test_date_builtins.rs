use super::helpers::compile_ok;

// ── date — format current timestamp ──────────────────────────

#[test]
fn date_format_current_timestamp() {
    compile_ok(
        r#"<?php
$formatted = date('Y-m-d');
echo strlen($formatted) === 10 ? 'ok' : 'bad length';
"#,
    );
}

// ── date with specific timestamp ──────────────────────────────

#[test]
fn date_with_specific_timestamp() {
    compile_ok(
        r#"<?php
$ts = 1704067200; // 2024-01-01 00:00:00 UTC
date_default_timezone_set('UTC');
echo date('Y', $ts);
echo date('m', $ts);
echo date('d', $ts);
"#,
    );
}

// ── mktime — create timestamp from components ─────────────────

#[test]
fn mktime_from_components() {
    compile_ok(
        r#"<?php
date_default_timezone_set('UTC');
$ts = mktime(0, 0, 0, 3, 15, 2024);
echo date('Y-m-d', $ts);
"#,
    );
}

// ── strtotime — parse date string ────────────────────────────

#[test]
fn strtotime_parse_iso() {
    compile_ok(
        r#"<?php
$ts = strtotime('2024-06-15');
echo $ts !== false ? 'parsed' : 'failed';
echo date('Y', $ts);
"#,
    );
}

// ── strtotime with relative dates ────────────────────────────

#[test]
fn strtotime_relative_expressions() {
    compile_ok(
        r#"<?php
date_default_timezone_set('UTC');
$base = strtotime('2024-01-01');
$next_week = strtotime('+1 week', $base);
echo date('d', $next_week);
$next_month = strtotime('+1 month', $base);
echo date('m', $next_month);
"#,
    );
}

// ── time — get current Unix timestamp ────────────────────────

#[test]
fn time_get_unix_timestamp() {
    compile_ok(
        r#"<?php
$t = time();
echo is_int($t) ? 'integer' : 'not integer';
echo $t > 1000000000 ? ':plausible' : ':implausible';
"#,
    );
}

// ── microtime — high resolution time ─────────────────────────

#[test]
fn microtime_string_form() {
    compile_ok(
        r#"<?php
$mt = microtime();
echo is_string($mt) ? 'string' : 'not string';
"#,
    );
}

// ── microtime(true) returns float ────────────────────────────

#[test]
fn microtime_float_form() {
    compile_ok(
        r#"<?php
$mt = microtime(true);
echo is_float($mt) ? 'float' : 'not float';
echo $mt > 1000000000.0 ? ':plausible' : ':implausible';
"#,
    );
}

// ── getdate — get date parts as array ────────────────────────

#[test]
fn getdate_returns_array() {
    compile_ok(
        r#"<?php
date_default_timezone_set('UTC');
$ts = mktime(10, 30, 0, 7, 4, 2024);
$info = getdate($ts);
echo $info['year'];
echo $info['mon'];
echo $info['mday'];
"#,
    );
}

// ── checkdate — validate Gregorian date ──────────────────────

#[test]
fn checkdate_validates_dates() {
    compile_ok(
        r#"<?php
echo checkdate(2, 29, 2024) ? 'valid' : 'invalid';
echo checkdate(2, 29, 2023) ? 'valid' : 'invalid';
echo checkdate(13, 1, 2024) ? 'valid' : 'invalid';
echo checkdate(12, 31, 9999) ? 'valid' : 'invalid';
"#,
    );
}

// ── date_create — create DateTime object ─────────────────────

#[test]
fn date_create_alias() {
    compile_ok(
        r#"<?php
$dt = date_create('2024-06-15');
echo $dt !== false ? 'created' : 'failed';
echo date_format($dt, 'Y');
"#,
    );
}

// ── date_format — format DateTime object ─────────────────────

#[test]
fn date_format_object() {
    compile_ok(
        r#"<?php
$dt = date_create('2024-12-25');
echo date_format($dt, 'Y-m-d');
echo date_format($dt, 'l');
"#,
    );
}

// ── date_modify — modify DateTime object ─────────────────────

#[test]
fn date_modify_object() {
    compile_ok(
        r#"<?php
$dt = date_create('2024-01-01');
date_modify($dt, '+3 months');
echo date_format($dt, 'Y-m-d');
"#,
    );
}

// ── date_diff — get DateInterval between two dates ────────────

#[test]
fn date_diff_two_objects() {
    compile_ok(
        r#"<?php
$d1 = date_create('2024-01-01');
$d2 = date_create('2024-06-15');
$diff = date_diff($d1, $d2);
echo $diff->days > 0 ? 'positive days' : 'zero or negative';
echo $diff->m > 0 ? ':has months' : ':no months';
"#,
    );
}

// ── date_add — add DateInterval to DateTime ───────────────────

#[test]
fn date_add_interval() {
    compile_ok(
        r#"<?php
$dt = date_create('2024-01-01');
$interval = new DateInterval('P2M');
date_add($dt, $interval);
echo date_format($dt, 'Y-m');
"#,
    );
}

// ── date_sub — subtract DateInterval from DateTime ───────────

#[test]
fn date_sub_interval() {
    compile_ok(
        r#"<?php
$dt = date_create('2024-06-15');
$interval = new DateInterval('P15D');
date_sub($dt, $interval);
echo date_format($dt, 'Y-m-d');
"#,
    );
}

// ── date_interval_create_from_date_string ────────────────────

#[test]
fn date_interval_from_date_string() {
    compile_ok(
        r#"<?php
$i = date_interval_create_from_date_string('3 weeks');
echo $i !== false ? 'created' : 'failed';
echo $i->days >= 0 ? ':has days' : ':no days';
"#,
    );
}

// ── date_timestamp_get — get Unix timestamp from DateTime ─────

#[test]
fn date_timestamp_get() {
    compile_ok(
        r#"<?php
date_default_timezone_set('UTC');
$dt = date_create('2024-01-01 00:00:00');
$ts = date_timestamp_get($dt);
echo is_int($ts) ? 'integer' : 'not integer';
echo date('Y', $ts);
"#,
    );
}

// ── date_timestamp_set — set Unix timestamp on DateTime ───────

#[test]
fn date_timestamp_set() {
    compile_ok(
        r#"<?php
date_default_timezone_set('UTC');
$dt = date_create('now');
$ts = mktime(0, 0, 0, 6, 15, 2024);
date_timestamp_set($dt, $ts);
echo date_format($dt, 'Y-m-d');
"#,
    );
}

// ── DateTime::createFromFormat static method ──────────────────

#[test]
fn datetime_create_from_format_static() {
    compile_ok(
        r#"<?php
$dt = DateTime::createFromFormat('d/m/Y', '25/12/2024');
echo $dt !== false ? 'created' : 'failed';
echo $dt->format('Y-m-d');
"#,
    );
}
