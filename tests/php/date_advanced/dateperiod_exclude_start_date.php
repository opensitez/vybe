<?php
// vybe-test: php/date_advanced/dateperiod_exclude_start_date
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable('2024-01-01');
$end = new DateTimeImmutable('2024-01-04');
$period = new DatePeriod($start, new DateInterval('P1D'), $end, DatePeriod::EXCLUDE_START_DATE);
$dates = [];
foreach ($period as $dt) { $dates[] = $dt->format('Y-m-d'); }
echo implode(',', $dates);
