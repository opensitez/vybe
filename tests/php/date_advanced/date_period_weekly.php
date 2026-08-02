<?php
// vybe-test: php/date_advanced/date_period_weekly
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable('2024-01-01');
$period = new DatePeriod($start, new DateInterval('P1W'), 4);
$weeks = [];
foreach ($period as $dt) { $weeks[] = $dt->format('W'); }
echo count($weeks);
