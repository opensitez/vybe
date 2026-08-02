<?php
// vybe-test: php/date_advanced/date_period_get_date_interval
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable('2024-01-01');
$interval = new DateInterval('P2W');
$period = new DatePeriod($start, $interval, 3);
echo $period->getDateInterval()->days >= 14 ? 'interval ok' : 'wrong';
