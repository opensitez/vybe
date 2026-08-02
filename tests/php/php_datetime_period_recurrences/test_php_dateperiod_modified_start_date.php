<?php
// vybe-test: php/php_datetime_period_recurrences/test_php_dateperiod_modified_start_date
// origin: languages/php/tests/php/test_php_datetime_period_recurrences.rs
// vybe-test-mode: compile

$start = (new DateTimeImmutable("2024-01-01"))->modify("+5 days");
$interval = new DateInterval("P1D");
$period = new DatePeriod($start, $interval, 2);

$first = null;
foreach ($period as $dt) {
    if ($first === null) $first = $dt;
}
echo $first->format("Y-m-d");
