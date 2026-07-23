use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: DateTime, DateTimeImmutable, DateInterval & DatePeriod — DateTimeImmutable, DateInterval, DatePeriod, format, modify, diff
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_datetime_immutable_chaining_and_immutability() {
    let out = run_prints(
        r#"<?php
$dt1 = new DateTimeImmutable("2024-01-01 12:00:00");
$dt2 = $dt1->add(new DateInterval("P10D"));

echo $dt1->format("Y-m-d") . " | " . $dt2->format("Y-m-d");
"#,
    );
    assert_eq!(out, vec!["2024-01-01 | 2024-01-11"]);
}

#[test]
fn test_php_datetime_create_from_format_parsing() {
    let out = run_prints(
        r#"<?php
$dateStr = "12/May/2024:15:30:45";
$dt = DateTimeImmutable::createFromFormat("d/M/Y:H:i:s", $dateStr);
echo $dt->format("Y-m-d H:i:s");
"#,
    );
    assert_eq!(out, vec!["2024-05-12 15:30:45"]);
}

#[test]
fn test_php_datetime_diff_interval_days() {
    let out = run_prints(
        r#"<?php
$d1 = new DateTimeImmutable("2024-01-01");
$d2 = new DateTimeImmutable("2024-01-15");
$interval = $d1->diff($d2);
echo $interval->format("%r%a days");
"#,
    );
    assert_eq!(out, vec!["14 days"]);
}

#[test]
fn test_php_dateperiod_iteration() {
    let out = run_prints(
        r#"<?php
$start = new DateTimeImmutable("2024-05-01");
$interval = new DateInterval("P1D");
$period = new DatePeriod($start, $interval, 2); // start + 2 occurrences

$dates = [];
foreach ($period as $dt) {
    $dates[] = $dt->format("Y-m-d");
}
echo implode(", ", $dates);
"#,
    );
    assert_eq!(out, vec!["2024-05-01, 2024-05-02, 2024-05-03"]);
}

#[test]
fn test_php_datetimezone_utc_offset_conversion() {
    let out = run_prints(
        r#"<?php
$dt = new DateTimeImmutable("2024-06-01 12:00:00", new DateTimeZone("UTC"));
$nyTz = new DateTimeZone("America/New_York");
$nyDt = $dt->setTimezone($nyTz);
echo $nyDt->format("Y-m-d H:i:s");
"#,
    );
    assert_eq!(out, vec!["2024-06-01 08:00:00"]);
}

#[test]
fn test_php_datetime_modify_relative_strings() {
    compile_ok(
        r#"<?php
$dt = new DateTimeImmutable("2024-01-01");
$nextMonday = $dt->modify("next monday");
$endOfMonth = $dt->modify("last day of this month");
echo $nextMonday->format("Y-m-d") . " " . $endOfMonth->format("Y-m-d");
"#,
    );
}

#[test]
fn test_php_dateinterval_create_from_datestring() {
    compile_ok(
        r#"<?php
$interval = DateInterval::createFromDateString("2 days 3 hours");
$dt = (new DateTimeImmutable("2024-01-01 00:00:00"))->add($interval);
echo $dt->format("Y-m-d H:i:s");
"#,
    );
}

#[test]
fn test_php_datetime_timestamp_get_set() {
    compile_ok(
        r#"<?php
$dt = new DateTimeImmutable();
$ts = $dt->getTimestamp();
$dt2 = $dt->setTimestamp($ts + 3600);
echo ($dt2->getTimestamp() - $ts);
"#,
    );
}

#[test]
fn test_php_datetime_comparison_operators() {
    compile_ok(
        r#"<?php
$dt1 = new DateTimeImmutable("2024-01-01");
$dt2 = new DateTimeImmutable("2024-06-01");
echo ($dt1 < $dt2 ? "EARLIER" : "LATER");
echo ($dt1 == new DateTimeImmutable("2024-01-01") ? "EQUAL" : "NOT_EQUAL");
"#,
    );
}

#[test]
fn test_php_dateperiod_exclude_start_date_option() {
    compile_ok(
        r#"<?php
$start = new DateTimeImmutable("2024-01-01");
$interval = new DateInterval("P1D");
$end = new DateTimeImmutable("2024-01-04");

$period = new DatePeriod($start, $interval, $end, DatePeriod::EXCLUDE_START_DATE);
foreach ($period as $dt) {
    echo $dt->format("Y-m-d") . "\n";
}
"#,
    );
}
