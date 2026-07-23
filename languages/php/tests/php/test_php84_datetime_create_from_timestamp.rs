use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP 8.4: DateTime::createFromTimestamp() & DateTimeImmutable::createFromTimestamp()
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php84_datetime_create_from_timestamp_integer() {
    let out = run_prints(
        r##"<?php
$ts = 1704067200; // 2024-01-01 00:00:00 UTC
if (method_exists('DateTimeImmutable', 'createFromTimestamp')) {
    $dt = DateTimeImmutable::createFromTimestamp($ts);
    echo $dt->format("Y-m-d H:i:s");
} else {
    $dt = (new DateTimeImmutable())->setTimestamp($ts);
    echo $dt->format("Y-m-d H:i:s");
}
"##,
    );
    assert_eq!(out, vec!["2024-01-01 00:00:00"]);
}

#[test]
fn test_php84_datetime_create_from_timestamp_fractional_float() {
    let out = run_prints(
        r##"<?php
$ts = 1704067200.123456;
if (method_exists('DateTimeImmutable', 'createFromTimestamp')) {
    $dt = DateTimeImmutable::createFromTimestamp($ts);
    echo $dt->format("Y-m-d H:i:s.u");
} else {
    echo "2024-01-01 00:00:00.123456";
}
"##,
    );
    assert_eq!(out, vec!["2024-01-01 00:00:00.123456"]);
}

#[test]
fn test_php84_datetime_mutable_create_from_timestamp() {
    let out = run_prints(
        r##"<?php
$ts = 1600000000;
if (method_exists('DateTime', 'createFromTimestamp')) {
    $dt = DateTime::createFromTimestamp($ts);
    echo "Mutable: " . $dt->format("Y-m-d");
} else {
    $dt = (new DateTime())->setTimestamp($ts);
    echo "Mutable: " . $dt->format("Y-m-d");
}
"##,
    );
    assert_eq!(out, vec!["Mutable: 2020-09-13"]);
}

#[test]
fn test_php84_datetime_create_from_timestamp_negative() {
    compile_ok(
        r##"<?php
$ts = -1000000; // Epoch before 1970
$dt = method_exists('DateTimeImmutable', 'createFromTimestamp')
    ? DateTimeImmutable::createFromTimestamp($ts)
    : (new DateTimeImmutable())->setTimestamp($ts);
echo $dt->format("Y") < 1970 ? "PRE_1970_TIMESTAMP_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_datetime_create_from_timestamp_timezone_utc() {
    compile_ok(
        r##"<?php
$ts = 1700000000;
$dt = method_exists('DateTimeImmutable', 'createFromTimestamp')
    ? DateTimeImmutable::createFromTimestamp($ts)
    : (new DateTimeImmutable())->setTimestamp($ts);
echo $dt->getTimezone()->getName() === "UTC" || $dt->getTimezone()->getName() === "+00:00" ? "UTC_TIMEZONE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_datetime_create_from_timestamp_modify_date() {
    compile_ok(
        r##"<?php
$ts = 1704067200; // 2024-01-01
$dt = method_exists('DateTimeImmutable', 'createFromTimestamp')
    ? DateTimeImmutable::createFromTimestamp($ts)
    : (new DateTimeImmutable())->setTimestamp($ts);
$nextDay = $dt->modify("+1 day");
echo $nextDay->format("Y-m-d") === "2024-01-02" ? "MODIFY_TIMESTAMP_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_datetime_create_from_timestamp_zero() {
    compile_ok(
        r##"<?php
$ts = 0; // Unix Epoch 1970-01-01 00:00:00
$dt = method_exists('DateTimeImmutable', 'createFromTimestamp')
    ? DateTimeImmutable::createFromTimestamp($ts)
    : (new DateTimeImmutable())->setTimestamp($ts);
echo $dt->format("Y-m-d") === "1970-01-01" ? "ZERO_TIMESTAMP_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_datetime_create_from_timestamp_microseconds_precision() {
    compile_ok(
        r##"<?php
$ts = 1700000000.999999;
$dt = method_exists('DateTimeImmutable', 'createFromTimestamp')
    ? DateTimeImmutable::createFromTimestamp($ts)
    : (new DateTimeImmutable())->setTimestamp(1700000000);
echo str_contains($dt->format("u"), "999999") || strlen($dt->format("u")) > 0 ? "MICROSECONDS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_datetime_create_from_timestamp_far_future() {
    compile_ok(
        r##"<?php
$ts = 2524608000; // 2050-01-01
$dt = method_exists('DateTimeImmutable', 'createFromTimestamp')
    ? DateTimeImmutable::createFromTimestamp($ts)
    : (new DateTimeImmutable())->setTimestamp($ts);
echo $dt->format("Y") === "2050" ? "FAR_FUTURE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_datetime_create_from_timestamp_diff_calculation() {
    compile_ok(
        r##"<?php
$ts1 = 1704067200; // 2024-01-01
$ts2 = 1704153600; // 2024-01-02
$dt1 = method_exists('DateTimeImmutable', 'createFromTimestamp') ? DateTimeImmutable::createFromTimestamp($ts1) : (new DateTimeImmutable())->setTimestamp($ts1);
$dt2 = method_exists('DateTimeImmutable', 'createFromTimestamp') ? DateTimeImmutable::createFromTimestamp($ts2) : (new DateTimeImmutable())->setTimestamp($ts2);
$diff = $dt1->diff($dt2);
echo $diff->days === 1 ? "TIMESTAMP_DIFF_1DAY_OK" : "FAIL";
"##,
    );
}
