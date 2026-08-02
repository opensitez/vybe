<?php
// vybe-test: php/date_advanced/date_period_recurrences
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable('2024-01-15');
$period = new DatePeriod($start, new DateInterval('P1M'), 5);
$months = [];
foreach ($period as $dt) { $months[] = $dt->format('m'); }
echo implode(',', $months);
