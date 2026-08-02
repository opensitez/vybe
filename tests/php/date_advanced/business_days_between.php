<?php
// vybe-test: php/date_advanced/business_days_between
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

function businessDays(DateTimeImmutable $start, DateTimeImmutable $end): int {
    $count = 0;
    $period = new DatePeriod($start, new DateInterval('P1D'), $end);
    foreach ($period as $day) {
        $dow = (int)$day->format('N'); // 1=Mon ... 7=Sun
        if ($dow < 6) $count++;
    }
    return $count;
}
$start = new DateTimeImmutable('2024-01-01');
$end   = new DateTimeImmutable('2024-01-08');
echo businessDays($start, $end);
