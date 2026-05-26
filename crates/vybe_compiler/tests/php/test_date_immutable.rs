use super::helpers::run_prints;

// ── DateTimeImmutable vs DateTime mutability ──────────────────

#[test] fn datetime_modify_mutates_original() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTime('2024-01-01');
$d->modify('+1 day');
echo $d->format('Y-m-d');
"#), vec!["2024-01-02"]);
}
#[test] fn datetimeimmutable_modify_returns_new_instance() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTimeImmutable('2024-01-01');
$d2 = $d->modify('+1 day');
echo $d->format('Y-m-d') . ',' . $d2->format('Y-m-d');
"#), vec!["2024-01-01,2024-01-02"]);
}
#[test] fn datetimeimmutable_setdate_returns_new() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTimeImmutable('2024-01-01');
$d2 = $d->setDate(2025, 6, 15);
echo $d->format('Y') . ',' . $d2->format('Y-m-d');
"#), vec!["2024,2025-06-15"]);
}
#[test] fn datetimeimmutable_settime_returns_new() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTimeImmutable('2024-01-01 00:00:00');
$d2 = $d->setTime(12, 30, 45);
echo $d2->format('H:i:s');
"#), vec!["12:30:45"]);
}
#[test] fn datetimeimmutable_add_returns_new() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTimeImmutable('2024-01-01');
$d2 = $d->add(new DateInterval('P7D'));
echo $d2->format('Y-m-d');
"#), vec!["2024-01-08"]);
}
#[test] fn datetimeimmutable_sub_returns_new() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTimeImmutable('2024-03-01');
$d2 = $d->sub(new DateInterval('P1M'));
echo $d2->format('Y-m-d');
"#), vec!["2024-02-01"]);
}

// ── DateTime formatting ───────────────────────────────────────

#[test] fn datetime_format_all_date_parts() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTime('2024-07-15');
echo $d->format('Y') . ',' . $d->format('m') . ',' . $d->format('d');
"#), vec!["2024,07,15"]);
}
#[test] fn datetime_format_day_of_week() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTime('2024-01-01');
echo $d->format('N');
"#), vec!["1"]);
}
#[test] fn datetime_format_week_number() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTime('2024-01-01');
echo $d->format('W');
"#), vec!["01"]);
}
#[test] fn datetime_format_timestamp() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTime('1970-01-01 00:00:00', new DateTimeZone('UTC'));
echo $d->getTimestamp();
"#), vec!["0"]);
}

// ── createFromFormat ──────────────────────────────────────────

#[test] fn datetime_create_from_custom_format() {
    assert_eq!(run_prints(r#"<?php
$d = DateTime::createFromFormat('d/m/Y', '25/12/2024');
echo $d->format('Y-m-d');
"#), vec!["2024-12-25"]);
}
#[test] fn datetime_create_from_format_with_time() {
    assert_eq!(run_prints(r#"<?php
$d = DateTime::createFromFormat('Y-m-d H:i', '2024-06-15 14:30');
echo $d->format('H:i');
"#), vec!["14:30"]);
}
#[test] fn datetimeimmutable_create_from_format() {
    assert_eq!(run_prints(r#"<?php
$d = DateTimeImmutable::createFromFormat('U', '0');
echo $d->setTimezone(new DateTimeZone('UTC'))->format('Y-m-d');
"#), vec!["1970-01-01"]);
}

// ── diff ──────────────────────────────────────────────────────

#[test] fn datetime_diff_days() {
    assert_eq!(run_prints(r#"<?php
$a = new DateTime('2024-01-01');
$b = new DateTime('2024-01-08');
echo $a->diff($b)->days;
"#), vec!["7"]);
}
#[test] fn datetime_diff_months() {
    assert_eq!(run_prints(r#"<?php
$a = new DateTime('2024-01-01');
$b = new DateTime('2024-04-01');
echo $a->diff($b)->m;
"#), vec!["3"]);
}
#[test] fn datetime_diff_negative_invert() {
    assert_eq!(run_prints(r#"<?php
$a = new DateTime('2024-01-08');
$b = new DateTime('2024-01-01');
$diff = $a->diff($b);
echo $diff->invert;
"#), vec!["1"]);
}

// ── Timezone ─────────────────────────────────────────────────

#[test] fn datetime_with_timezone() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTime('2024-01-01 00:00:00', new DateTimeZone('UTC'));
echo $d->getTimezone()->getName();
"#), vec!["UTC"]);
}
#[test] fn datetime_settimezone_converts() {
    assert_eq!(run_prints(r#"<?php
$d = new DateTime('2024-01-01 00:00:00', new DateTimeZone('UTC'));
$d->setTimezone(new DateTimeZone('America/New_York'));
echo $d->format('P');
"#), vec!["-05:00"]);
}

// ── strtotime ────────────────────────────────────────────────

#[test] fn strtotime_relative_next_monday() {
    assert_eq!(run_prints(r#"<?php
$ts = strtotime('next monday', mktime(0,0,0,1,1,2024));
echo date('N', $ts);
"#), vec!["1"]);
}
#[test] fn strtotime_plus_days() {
    assert_eq!(run_prints(r#"<?php
$base = mktime(0,0,0,1,1,2024);
$ts = strtotime('+7 days', $base);
echo date('d', $ts);
"#), vec!["08"]);
}
#[test] fn strtotime_invalid_returns_false() {
    assert_eq!(run_prints(r#"<?php echo var_export(strtotime('not a date'), true); "#), vec!["false"]);
}

// ── checkdate ────────────────────────────────────────────────

#[test] fn checkdate_valid_date() {
    assert_eq!(run_prints(r#"<?php echo checkdate(2, 29, 2024) ? 'yes' : 'no'; "#), vec!["yes"]);
}
#[test] fn checkdate_invalid_feb29_non_leap() {
    assert_eq!(run_prints(r#"<?php echo checkdate(2, 29, 2023) ? 'yes' : 'no'; "#), vec!["no"]);
}
#[test] fn checkdate_invalid_month() {
    assert_eq!(run_prints(r#"<?php echo checkdate(13, 1, 2024) ? 'yes' : 'no'; "#), vec!["no"]);
}
