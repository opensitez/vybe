<?php
// vybe-test: php/date_advanced/date_interval_invert
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$a = new DateTimeImmutable('2024-03-01');
$b = new DateTimeImmutable('2024-01-01');
$diff = $a->diff($b);
echo $diff->invert;  // 1 (b < a, so diff is negative direction)
echo ':' . $diff->m;
