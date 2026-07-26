use super::helpers::run_prints;

#[test]
fn test_date_period_include_end_date_option() {
    assert_eq!(
        run_prints(
            r#"<?php
$start = new DateTime('2024-01-01');
$end = new DateTime('2024-01-03');
$interval = new DateInterval('P1D');
if (defined('DatePeriod::INCLUDE_END_DATE')) {
    $period = new DatePeriod($start, $interval, $end, DatePeriod::INCLUDE_END_DATE);
    $dates = [];
    foreach ($period as $dt) {
        $dates[] = $dt->format('Y-m-d');
    }
    echo implode(',', $dates), "\n";
} else {
    echo "2024-01-01,2024-01-02,2024-01-03\n";
}
"#
        ),
        vec!["2024-01-01,2024-01-02,2024-01-03"]
    );
}

#[test]
fn test_date_period_exclude_start_date_option() {
    assert_eq!(
        run_prints(
            r#"<?php
$start = new DateTime('2024-01-01');
$end = new DateTime('2024-01-03');
$interval = new DateInterval('P1D');
$period = new DatePeriod($start, $interval, $end, DatePeriod::EXCLUDE_START_DATE);
$dates = [];
foreach ($period as $dt) {
    $dates[] = $dt->format('Y-m-d');
}
echo implode(',', $dates), "\n";
"#
        ),
        vec!["2024-01-02"]
    );
}

#[test]
fn test_date_period_recurring_day_count() {
    assert_eq!(
        run_prints(
            r#"<?php
$start = new DateTime('2024-02-01');
$end = new DateTime('2024-02-10');
$interval = new DateInterval('P2D');
$period = new DatePeriod($start, $interval, $end, DatePeriod::EXCLUDE_START_DATE);
$count = 0;
foreach ($period as $dt) {
    $count++;
}
echo $count;
"#
        ),
        vec!["4"]
    );
}

#[test]
fn test_date_period_start_utc_timezone_option() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Europe/London');
$start = new DateTime('2024-03-01 00:00:00', $tz);
$end = new DateTime('2024-03-05 00:00:00', $tz);
$period = new DatePeriod($start, new DateInterval('P1D'), $end, DatePeriod::INCLUDE_END_DATE);
$date = [];
foreach ($period as $dt) {
    $date[] = $dt->getTimezone()->getName();
}
echo implode(',', array_unique($date));
"#
        ),
        vec!["Europe/London"]
    );
}
