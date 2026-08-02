<?php
// vybe-test: php/date_advanced/date_period_include_start_date
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable('2024-01-01');
$end   = new DateTimeImmutable('2024-04-01');
// With DatePeriod::INCLUDE_START_DATE (default) vs EXCLUDE_START_DATE
$incl = new DatePeriod($start, new DateInterval('P1M'), $end);
$count = 0;
foreach ($incl as $dt) { $count++; }
echo $count;  // 3
