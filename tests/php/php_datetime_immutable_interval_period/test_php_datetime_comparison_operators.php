<?php
// vybe-test: php/php_datetime_immutable_interval_period/test_php_datetime_comparison_operators
// origin: languages/php/tests/php/test_php_datetime_immutable_interval_period.rs
// vybe-test-mode: compile

$dt1 = new DateTimeImmutable("2024-01-01");
$dt2 = new DateTimeImmutable("2024-06-01");
echo ($dt1 < $dt2 ? "EARLIER" : "LATER");
echo ($dt1 == new DateTimeImmutable("2024-01-01") ? "EQUAL" : "NOT_EQUAL");
