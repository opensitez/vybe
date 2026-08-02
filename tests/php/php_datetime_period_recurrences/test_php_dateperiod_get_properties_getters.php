<?php
// vybe-test: php/php_datetime_period_recurrences/test_php_dateperiod_get_properties_getters
// origin: languages/php/tests/php/test_php_datetime_period_recurrences.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable("2024-01-01");
$interval = new DateInterval("P1W");
$end = new DateTimeImmutable("2024-02-01");
$period = new DatePeriod($start, $interval, $end);

echo "Start=" . $period->getStartDate()->format("Y-m-d") . " End=" . $period->getEndDate()->format("Y-m-d");
