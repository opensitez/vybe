<?php
// vybe-test: php/php_datetime_period_recurrences/test_php_dateperiod_with_timezone_context
// origin: languages/php/tests/php/test_php_datetime_period_recurrences.rs
// vybe-test-mode: compile

$tz = new DateTimeZone("America/New_York");
$start = new DateTimeImmutable("2024-01-01", $tz);
$interval = new DateInterval("P1M");
$end = new DateTimeImmutable("2024-03-01", $tz);
$period = new DatePeriod($start, $interval, $end);

foreach ($period as $dt) {
    echo $dt->getTimezone()->getName() . " ";
}
