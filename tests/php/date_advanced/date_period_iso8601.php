<?php
// vybe-test: php/date_advanced/date_period_iso8601
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$period = new DatePeriod('R3/2024-01-01T00:00:00Z/P1D');
$days = [];
foreach ($period as $dt) { $days[] = $dt->format('Y-m-d'); }
echo implode('|', $days);
