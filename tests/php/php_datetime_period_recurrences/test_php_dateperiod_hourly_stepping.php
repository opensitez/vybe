<?php
// vybe-test: php/php_datetime_period_recurrences/test_php_dateperiod_hourly_stepping
// origin: languages/php/tests/php/test_php_datetime_period_recurrences.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable("2024-05-12 00:00:00");
$interval = new DateInterval("PT4H"); // every 4 hours
$end = new DateTimeImmutable("2024-05-12 12:00:00");
$period = new DatePeriod($start, $interval, $end);

$steps = 0;
foreach ($period as $dt) { $steps++; }
echo "Hourly steps: $steps";
