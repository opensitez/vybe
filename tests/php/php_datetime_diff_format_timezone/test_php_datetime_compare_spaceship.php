<?php
// vybe-test: php/php_datetime_diff_format_timezone/test_php_datetime_compare_spaceship
// origin: languages/php/tests/php/test_php_datetime_diff_format_timezone.rs
// vybe-test-mode: compile

$dt1 = new DateTimeImmutable("2024-01-01");
$dt2 = new DateTimeImmutable("2024-05-01");
echo ($dt1 <=> $dt2) < 0 ? "EARLIER" : "LATER";
