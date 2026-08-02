<?php
// vybe-test: php/date_advanced/datetime_immutable_modify
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$dt = new DateTimeImmutable('2024-01-15');
$next_month = $dt->modify('+1 month');
echo $dt->format('Y-m-d');          // unchanged
echo ':' . $next_month->format('Y-m-d');
