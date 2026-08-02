<?php
// vybe-test: php/date_advanced/date_period_daily
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable('2024-03-01');
$end   = new DateTimeImmutable('2024-03-08');
$period = new DatePeriod($start, new DateInterval('P1D'), $end);
$dates = [];
foreach ($period as $dt) { $dates[] = $dt->format('d'); }
echo implode(',', $dates);
