<?php
// vybe-test: php/date_advanced/recurring_monthly_dates
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$start = new DateTimeImmutable('2024-01-31');
$dates = [];
for ($i = 0; $i < 4; $i++) {
    $dates[] = $start->modify("+$i month")->format('Y-m-d');
}
echo implode('|', $dates);
echo ':' . count($dates);
