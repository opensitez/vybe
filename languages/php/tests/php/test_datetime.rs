use super::helpers::run_prints;

// ── date() / time() basics ───────────────────────────────────────
#[test]
fn time_returns_integer() {
    assert_eq!(
        run_prints(
            r#"<?php
$t = time();
echo is_int($t) ? "yes" : "no";
"#
        ),
        &["yes"]
    );
}

#[test]
fn date_format_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = mktime(14, 30, 0, 6, 15, 2024);
echo date("Y-m-d", $ts);
echo date("H:i:s", $ts);
"#
        ),
        // echo emits no newline, so PHP concatenates the two into one line.
        &["2024-06-1514:30:00"]
    );
}

#[test]
fn date_format_parts() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = mktime(0, 0, 0, 12, 25, 2023);
echo date("Y", $ts);
echo date("m", $ts);
echo date("d", $ts);
echo date("D", $ts);
"#
        ),
        // echo emits no newline: all four parts concatenate into one line.
        &["20231225Mon"]
    );
}

#[test]
fn mktime_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts1 = mktime(0, 0, 0, 1, 1, 2020);
$ts2 = mktime(0, 0, 0, 1, 2, 2020);
echo $ts2 - $ts1;
"#
        ),
        &["86400"]
    );
}

#[test]
fn date_default_timezone_set_utc_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo date_default_timezone_set('UTC') ? 'yes' : 'no';
"#
        ),
        &["yes"]
    );
}

// ── strtotime ────────────────────────────────────────────────────
#[test]
fn strtotime_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = strtotime("2024-01-15");
echo date("Y-m-d", $ts);
"#
        ),
        &["2024-01-15"]
    );
}

#[test]
fn strtotime_relative() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = strtotime("2024-06-15");
$next = strtotime("+7 days", $base);
echo date("Y-m-d", $next);
$prev = strtotime("-1 month", $base);
echo date("Y-m-d", $prev);
"#
        ),
        // echo emits no newline: both dates concatenate into one line.
        &["2024-06-222024-05-15"]
    );
}

// ── checkdate ────────────────────────────────────────────────────
#[test]
fn checkdate_valid() {
    assert_eq!(
        run_prints(
            r#"<?php
echo checkdate(2, 29, 2024) ? "valid" : "invalid";
echo checkdate(2, 29, 2023) ? "valid" : "invalid";
echo checkdate(13, 1, 2024) ? "valid" : "invalid";
echo checkdate(12, 31, 2024) ? "valid" : "invalid";
"#
        ),
        &["validinvalidinvalidvalid"]
    );
}

// ── DateTime class ───────────────────────────────────────────────
#[test]
fn datetime_construct() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-06-15 14:30:00");
echo $dt->format("Y-m-d");
echo $dt->format("H:i:s");
"#
        ),
        &["2024-06-1514:30:00"]
    );
}

#[test]
fn datetime_modify() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-01-01");
$dt->modify("+6 months");
echo $dt->format("Y-m-d");
$dt->modify("-10 days");
echo $dt->format("Y-m-d");
"#
        ),
        &["2024-07-012024-06-21"]
    );
}

#[test]
fn datetime_diff() {
    assert_eq!(
        run_prints(
            r#"<?php
$d1 = new DateTime("2024-01-01");
$d2 = new DateTime("2024-03-01");
$diff = $d1->diff($d2);
echo $diff->days;
echo $diff->m;
"#
        ),
        &["602"]
    );
}

#[test]
fn datetime_format_various() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-12-25 09:30:45");
echo $dt->format("l");
echo $dt->format("F j, Y");
echo $dt->format("g:i A");
"#
        ),
        &["WednesdayDecember 25, 20249:30 AM"]
    );
}

#[test]
fn datetime_create_from_format() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTime::createFromFormat("d/m/Y", "25/12/2024");
echo $dt->format("Y-m-d");
"#
        ),
        &["2024-12-25"]
    );
}

#[test]
fn datetime_get_timestamp() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-01-01 00:00:00");
$ts = $dt->getTimestamp();
echo is_int($ts) ? "yes" : "no";
echo date("Y", $ts);
"#
        ),
        &["yes2024"]
    );
}

// ── DateInterval ─────────────────────────────────────────────────
#[test]
fn dateinterval_construct() {
    assert_eq!(
        run_prints(
            r#"<?php
$interval = new DateInterval("P1Y2M3D");
echo $interval->y;
echo $interval->m;
echo $interval->d;
"#
        ),
        &["123"]
    );
}

#[test]
fn datetime_add_interval() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-01-01");
$interval = new DateInterval("P30D");
$dt->add($interval);
echo $dt->format("Y-m-d");
"#
        ),
        &["2024-01-31"]
    );
}

#[test]
fn datetime_sub_interval() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-06-15");
$interval = new DateInterval("P3M");
$dt->sub($interval);
echo $dt->format("Y-m-d");
"#
        ),
        &["2024-03-15"]
    );
}

// ── getdate / localtime ──────────────────────────────────────────
#[test]
fn getdate_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = mktime(14, 30, 0, 6, 15, 2024);
$info = getdate($ts);
echo $info["year"];
echo $info["mon"];
echo $info["mday"];
echo $info["hours"];
"#
        ),
        &["202461514"]
    );
}

// ── Date arithmetic patterns ─────────────────────────────────────
#[test]
fn days_between_dates() {
    assert_eq!(
        run_prints(
            r#"<?php
$start = new DateTime("2024-01-01");
$end = new DateTime("2024-12-31");
$diff = $start->diff($end);
echo $diff->days;
"#
        ),
        &["365"]
    );
}

#[test]
fn date_comparison() {
    assert_eq!(
        run_prints(
            r#"<?php
$d1 = new DateTime("2024-01-01");
$d2 = new DateTime("2024-06-15");
echo $d1 < $d2 ? "before" : "after";
echo $d1 == $d2 ? "equal" : "not equal";
"#
        ),
        &["beforenot equal"]
    );
}

#[test]
fn date_immutable() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTimeImmutable("2024-01-01");
$modified = $dt->modify("+1 month");
echo $dt->format("Y-m-d");
echo $modified->format("Y-m-d");
"#
        ),
        // echo emits no newline: both dates concatenate into one line.
        &["2024-01-012024-02-01"]
    );
}
