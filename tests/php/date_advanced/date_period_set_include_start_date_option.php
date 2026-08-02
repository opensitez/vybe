<?php
// vybe-test: php/date_advanced/date_period_set_include_start_date_option
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable('2024-01-01');
$period = new DatePeriod($start, new DateInterval('P1D'), 3, DatePeriod::INCLUDE_END_DATE);
$count = 0;
foreach ($period as $dt) { $count++; }
echo $count;
