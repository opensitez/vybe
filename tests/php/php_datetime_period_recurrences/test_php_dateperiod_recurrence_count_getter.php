<?php
// vybe-test: php/php_datetime_period_recurrences/test_php_dateperiod_recurrence_count_getter
// origin: languages/php/tests/php/test_php_datetime_period_recurrences.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable("2024-01-01");
$interval = new DateInterval("P1D");
$period = new DatePeriod($start, $interval, 5);
echo "Recurrences: " . $period->getRecurrences();
