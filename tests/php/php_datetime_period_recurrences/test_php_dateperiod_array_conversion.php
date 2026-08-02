<?php
// vybe-test: php/php_datetime_period_recurrences/test_php_dateperiod_array_conversion
// origin: languages/php/tests/php/test_php_datetime_period_recurrences.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable("2024-01-01");
$interval = new DateInterval("P1D");
$period = new DatePeriod($start, $interval, 3);
$arr = iterator_to_array($period);
echo count($arr) === 4 ? "ITERATOR_TO_ARRAY_OK" : "FAIL";
