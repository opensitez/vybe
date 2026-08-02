<?php
// vybe-test: php/php_datetime_immutable_interval_period/test_php_datetime_modify_relative_strings
// origin: languages/php/tests/php/test_php_datetime_immutable_interval_period.rs
// vybe-test-mode: compile

$dt = new DateTimeImmutable("2024-01-01");
$nextMonday = $dt->modify("next monday");
$endOfMonth = $dt->modify("last day of this month");
echo $nextMonday->format("Y-m-d") . " " . $endOfMonth->format("Y-m-d");
