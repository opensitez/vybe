<?php
// vybe-test: php/date_advanced/month_boundaries
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$months = [];
$start = new DateTimeImmutable('2024-01-01');
$period = new DatePeriod($start, new DateInterval('P1M'), 12);
foreach ($period as $dt) {
    $months[] = $dt->format('Y-m');
}
echo count($months) . ':' . $months[0] . ':' . end($months);
