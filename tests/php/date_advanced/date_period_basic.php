<?php
// vybe-test: php/date_advanced/date_period_basic
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$start    = new DateTimeImmutable('2024-01-01');
$interval = new DateInterval('P1M');  // 1 month
$end      = new DateTimeImmutable('2024-06-01');
$period   = new DatePeriod($start, $interval, $end);
$count = 0;
foreach ($period as $dt) { $count++; }
echo $count;  // 5 (Jan through May)
