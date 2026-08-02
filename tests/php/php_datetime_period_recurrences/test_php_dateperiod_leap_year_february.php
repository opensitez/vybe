<?php
// vybe-test: php/php_datetime_period_recurrences/test_php_dateperiod_leap_year_february
// origin: languages/php/tests/php/test_php_datetime_period_recurrences.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable("2024-02-27");
$interval = new DateInterval("P1D");
$end = new DateTimeImmutable("2024-03-02");
$period = new DatePeriod($start, $interval, $end);

$dates = [];
foreach ($period as $dt) { $dates[] = $dt->format("m-d"); }
echo implode(",", $dates); // 02-27, 02-28, 02-29, 03-01
