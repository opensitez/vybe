use super::helpers::run_prints;

#[test]
fn test_date_period_iso_string_recurrence() {
    assert_eq!(
        run_prints(
            r#"<?php
$period = new DatePeriod('R3/2024-01-01T00:00:00Z/P1D');
$dates = [];
foreach ($period as $dt) {
    $dates[] = $dt->format('Y-m-d');
}
echo implode(',', $dates), "\n";
"#
        ),
        vec!["2024-01-01,2024-01-02,2024-01-03,2024-01-04"]
    );
}

#[test]
fn test_date_period_get_recurrences() {
    assert_eq!(
        run_prints(
            r#"<?php
$start = new DateTime('2024-01-01');
$interval = new DateInterval('P1D');
$period = new DatePeriod($start, $interval, 5);
echo $period->getRecurrences(), "\n";
"#
        ),
        vec!["5"]
    );
}

#[test]
fn test_date_period_iso_exclude_start() {
    assert_eq!(
        run_prints(
        r#"<?php
$period = new DatePeriod('R3/2024-01-01T00:00:00Z/P1D');
$dates = [];
foreach ($period as $dt) {
    $dates[] = $dt->format('Y-m-d');
}
echo implode('|', $dates), "\n";
"#
        ),
        vec!["2024-01-01|2024-01-02|2024-01-03|2024-01-04"]
    );
}

#[test]
fn test_date_period_iso_get_start_end_dates() {
    assert_eq!(
        run_prints(
        r#"<?php
$period = new DatePeriod('R2/2024-12-30T00:00:00Z/P1D');
$end = $period->getEndDate();
echo $period->getStartDate()->format('Y-m-d');
echo '|';
echo $end instanceof DateTimeInterface ? $end->format('Y-m-d') : 'none';
"#
        ),
        vec!["2024-12-30|2024-12-31"]
    );
}
