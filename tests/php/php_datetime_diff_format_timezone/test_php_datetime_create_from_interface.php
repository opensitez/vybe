<?php
// vybe-test: php/php_datetime_diff_format_timezone/test_php_datetime_create_from_interface
// origin: languages/php/tests/php/test_php_datetime_diff_format_timezone.rs
// vybe-test-mode: compile

$dt1 = new DateTime("2024-01-01");
$dt2 = DateTimeImmutable::createFromInterface($dt1);
echo $dt2->format("Y-m-d");
