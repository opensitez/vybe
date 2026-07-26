use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: DatePeriod Recurrences & ISO Strings — DatePeriod, ISO8601 recurrences, DateInterval stepping, include_start_date
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_dateperiod_iteration_start_end_interval() {
    let out = run_prints(
        r#"<?php
$start = new DateTimeImmutable("2024-01-01");
$interval = new DateInterval("P1M"); // 1 month
$end = new DateTimeImmutable("2024-04-01");

$period = new DatePeriod($start, $interval, $end);
$dates = [];
foreach ($period as $dt) {
    $dates[] = $dt->format("Y-m");
}
echo implode(", ", $dates);
"#,
    );
    assert_eq!(out, vec!["2024-01, 2024-02, 2024-03"]);
}

#[test]
fn test_php_dateperiod_iso8601_string_constructor() {
    let out = run_prints(
        r#"<?php
// Repeat 3 times every 2 days starting from 2024-05-01
$period = new DatePeriod("R3/2024-05-01T00:00:00Z/P2D");
$dates = [];
foreach ($period as $dt) {
    $dates[] = $dt->format("Y-m-d");
}
echo implode(", ", $dates);
"#,
    );
    assert_eq!(out, vec!["2024-05-01, 2024-05-03, 2024-05-05, 2024-05-07"]);
}

#[test]
fn test_php_dateperiod_exclude_start_date_option() {
    let out = run_prints(
        r#"<?php
$start = new DateTimeImmutable("2024-01-01");
$interval = new DateInterval("P1D");
$recurrences = 2;

$period = new DatePeriod($start, $interval, $recurrences, DatePeriod::EXCLUDE_START_DATE);
$dates = [];
foreach ($period as $dt) {
    $dates[] = $dt->format("Y-m-d");
}
echo implode(", ", $dates);
"#,
    );
    assert_eq!(out, vec!["2024-01-02, 2024-01-03"]);
}

#[test]
fn test_php_dateperiod_get_properties_getters() {
    compile_ok(
        r#"<?php
$start = new DateTimeImmutable("2024-01-01");
$interval = new DateInterval("P1W");
$end = new DateTimeImmutable("2024-02-01");
$period = new DatePeriod($start, $interval, $end);

echo "Start=" . $period->getStartDate()->format("Y-m-d") . " End=" . $period->getEndDate()->format("Y-m-d");
"#,
    );
}

#[test]
fn test_php_dateperiod_recurrence_count_getter() {
    compile_ok(
        r#"<?php
$start = new DateTimeImmutable("2024-01-01");
$interval = new DateInterval("P1D");
$period = new DatePeriod($start, $interval, 5);
echo "Recurrences: " . $period->getRecurrences();
"#,
    );
}

#[test]
fn test_php_dateperiod_recurrence_count_zero() {
    let out = run_prints(
        r#"<?php
$start = new DateTimeImmutable('2024-01-01');
$interval = new DateInterval('P1D');
$period = new DatePeriod($start, $interval, 0);
echo $period->getRecurrences();
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_php_dateperiod_hourly_stepping_counted() {
    let out = run_prints(
        r#"<?php
$start = new DateTimeImmutable('2024-05-12 00:00:00');
$interval = new DateInterval('PT6H');
$end = new DateTimeImmutable('2024-05-13 00:00:00');
$period = new DatePeriod($start, $interval, $end);
$count = 0;
foreach ($period as $_dt) { $count++; }
echo $count;
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_php_dateperiod_with_include_end_date() {
    let out = run_prints(
        r#"<?php
$start = new DateTimeImmutable('2024-01-01');
$interval = new DateInterval('P1D');
$end = new DateTimeImmutable('2024-01-04');
$period = new DatePeriod($start, $interval, $end, DatePeriod::INCLUDE_END_DATE);
$dates = [];
foreach ($period as $dt) { $dates[] = $dt->format('Y-m-d'); }
echo implode('|', $dates);
"#,
    );
    assert_eq!(out, vec!["2024-01-01|2024-01-02|2024-01-03|2024-01-04"]);
}

#[test]
fn test_php_dateperiod_timezone_offset_preserved() {
    let out = run_prints(
        r#"<?php
$tz = new DateTimeZone('Europe/Paris');
$start = new DateTimeImmutable('2024-01-01', $tz);
$period = new DatePeriod($start, new DateInterval('P1D'), 2);
$tzName = null;
foreach ($period as $dt) {
    $tzName = $dt->getTimezone()->getName();
    break;
}
echo $tzName;
"#,
    );
    assert_eq!(out, vec!["Europe/Paris"]);
}

#[test]
fn test_php_dateperiod_hourly_stepping() {
    compile_ok(
        r#"<?php
$start = new DateTimeImmutable("2024-05-12 00:00:00");
$interval = new DateInterval("PT4H"); // every 4 hours
$end = new DateTimeImmutable("2024-05-12 12:00:00");
$period = new DatePeriod($start, $interval, $end);

$steps = 0;
foreach ($period as $dt) { $steps++; }
echo "Hourly steps: $steps";
"#,
    );
}

#[test]
fn test_php_dateperiod_with_timezone_context() {
    compile_ok(
        r#"<?php
$tz = new DateTimeZone("America/New_York");
$start = new DateTimeImmutable("2024-01-01", $tz);
$interval = new DateInterval("P1M");
$end = new DateTimeImmutable("2024-03-01", $tz);
$period = new DatePeriod($start, $interval, $end);

foreach ($period as $dt) {
    echo $dt->getTimezone()->getName() . " ";
}
"#,
    );
}

#[test]
fn test_php_dateperiod_modified_start_date() {
    compile_ok(
        r#"<?php
$start = (new DateTimeImmutable("2024-01-01"))->modify("+5 days");
$interval = new DateInterval("P1D");
$period = new DatePeriod($start, $interval, 2);

$first = null;
foreach ($period as $dt) {
    if ($first === null) $first = $dt;
}
echo $first->format("Y-m-d");
"#,
    );
}

#[test]
fn test_php_dateperiod_array_conversion() {
    compile_ok(
        r#"<?php
$start = new DateTimeImmutable("2024-01-01");
$interval = new DateInterval("P1D");
$period = new DatePeriod($start, $interval, 3);
$arr = iterator_to_array($period);
echo count($arr) === 4 ? "ITERATOR_TO_ARRAY_OK" : "FAIL";
"#,
    );
}

#[test]
fn test_php_dateperiod_leap_year_february() {
    compile_ok(
        r#"<?php
$start = new DateTimeImmutable("2024-02-27");
$interval = new DateInterval("P1D");
$end = new DateTimeImmutable("2024-03-02");
$period = new DatePeriod($start, $interval, $end);

$dates = [];
foreach ($period as $dt) { $dates[] = $dt->format("m-d"); }
echo implode(",", $dates); // 02-27, 02-28, 02-29, 03-01
"#,
    );
}
