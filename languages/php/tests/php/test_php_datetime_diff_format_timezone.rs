use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: DateTime Calculations & Formatting Specifiers — format specifiers, DateInterval format, timezone offsets, diff
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_datetime_format_specifiers_full() {
    let out = run_prints(
        r#"<?php
$dt = new DateTimeImmutable("2024-05-12 14:30:45", new DateTimeZone("UTC"));
echo $dt->format("Y-m-d H:i:s P");
"#,
    );
    assert_eq!(out, vec!["2024-05-12 14:30:45 +00:00"]);
}

#[test]
fn test_php_dateinterval_format_specifiers() {
    let out = run_prints(
        r#"<?php
$d1 = new DateTimeImmutable("2024-01-01 00:00:00");
$d2 = new DateTimeImmutable("2024-03-15 10:30:00");
$diff = $d1->diff($d2);

echo $diff->format("%m months %d days %h hours %i minutes");
"#,
    );
    assert_eq!(out, vec!["2 months 14 days 10 hours 30 minutes"]);
}

#[test]
fn test_php_datetime_modify_relative_business_days() {
    let out = run_prints(
        r#"<?php
$dt = new DateTimeImmutable("2024-05-10"); // Friday
$monday = $dt->modify("+3 days");
echo $monday->format("Y-m-d");
"#,
    );
    assert_eq!(out, vec!["2024-05-13"]);
}

#[test]
fn test_php_datetime_sub_interval() {
    let out = run_prints(
        r#"<?php
$dt = new DateTimeImmutable("2024-05-12");
$interval = new DateInterval("P1W"); // 1 week
$earlier = $dt->sub($interval);
echo $earlier->format("Y-m-d");
"#,
    );
    assert_eq!(out, vec!["2024-05-05"]);
}

#[test]
fn test_php_datetime_set_time_and_date() {
    compile_ok(
        r#"<?php
$dt = (new DateTimeImmutable())
    ->setDate(2025, 12, 25)
    ->setTime(10, 0, 0, 500000);

echo $dt->format("Y-m-d H:i:s.u");
"#,
    );
}

#[test]
fn test_php_datetime_create_from_interface() {
    compile_ok(
        r#"<?php
$dt1 = new DateTime("2024-01-01");
$dt2 = DateTimeImmutable::createFromInterface($dt1);
echo $dt2->format("Y-m-d");
"#,
    );
}

#[test]
fn test_php_datetime_isostring_formatting() {
    compile_ok(
        r#"<?php
$dt = new DateTimeImmutable("2024-05-12T15:30:00+02:00");
echo $dt->format(DateTimeInterface::ATOM) . " | " . $dt->format(DateTimeInterface::ISO8601);
"#,
    );
}

#[test]
fn test_php_dateinterval_inverted_flag() {
    compile_ok(
        r#"<?php
$d1 = new DateTimeImmutable("2024-05-12");
$d2 = new DateTimeImmutable("2024-01-01");
$diff = $d1->diff($d2);
echo $diff->invert ? "INVERTED_PAST" : "FUTURE";
"#,
    );
}

#[test]
fn test_php_datetime_microtime_precision() {
    compile_ok(
        r#"<?php
$dt = DateTimeImmutable::createFromFormat("U.u", sprintf("%.6f", microtime(true)));
echo $dt->format("Y-m-d H:i:s.u");
"#,
    );
}

#[test]
fn test_php_datetime_compare_spaceship() {
    compile_ok(
        r#"<?php
$dt1 = new DateTimeImmutable("2024-01-01");
$dt2 = new DateTimeImmutable("2024-05-01");
echo ($dt1 <=> $dt2) < 0 ? "EARLIER" : "LATER";
"#,
    );
}
