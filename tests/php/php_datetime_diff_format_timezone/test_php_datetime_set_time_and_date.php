<?php
// vybe-test: php/php_datetime_diff_format_timezone/test_php_datetime_set_time_and_date
// origin: languages/php/tests/php/test_php_datetime_diff_format_timezone.rs
// vybe-test-mode: compile

$dt = (new DateTimeImmutable())
    ->setDate(2025, 12, 25)
    ->setTime(10, 0, 0, 500000);

echo $dt->format("Y-m-d H:i:s.u");
