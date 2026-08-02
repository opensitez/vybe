<?php
// vybe-test: php/date_advanced/datetime_comparison
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$a = new DateTimeImmutable('2024-01-01');
$b = new DateTimeImmutable('2024-06-15');
echo ($a < $b)  ? 'a before b' : 'a not before b';
echo ($a == $b) ? ':equal'     : ':not equal';
