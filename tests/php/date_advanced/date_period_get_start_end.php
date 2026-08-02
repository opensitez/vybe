<?php
// vybe-test: php/date_advanced/date_period_get_start_end
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable('2024-01-01');
$end   = new DateTimeImmutable('2024-12-31');
$period = new DatePeriod($start, new DateInterval('P1M'), $end);
echo $period->getStartDate()->format('Y-m-d');
echo ':' . $period->getEndDate()->format('Y-m-d');
