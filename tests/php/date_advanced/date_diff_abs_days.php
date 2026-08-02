<?php
// vybe-test: php/date_advanced/date_diff_abs_days
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$a = new DateTimeImmutable('2024-01-01');
$b = new DateTimeImmutable('2024-12-31');
$diff = $a->diff($b);
echo $diff->days;  // 365 (2024 is leap year)
