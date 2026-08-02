<?php
// vybe-test: php/php_datetime_immutable_interval_period/test_php_dateperiod_exclude_start_date_option
// origin: languages/php/tests/php/test_php_datetime_immutable_interval_period.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable("2024-01-01");
$interval = new DateInterval("P1D");
$end = new DateTimeImmutable("2024-01-04");

$period = new DatePeriod($start, $interval, $end, DatePeriod::EXCLUDE_START_DATE);
foreach ($period as $dt) {
    echo $dt->format("Y-m-d") . "\n";
}
