<?php
// vybe-test: php/php_datetime_diff_format_timezone/test_php_datetime_isostring_formatting
// origin: languages/php/tests/php/test_php_datetime_diff_format_timezone.rs
// vybe-test-mode: compile

$dt = new DateTimeImmutable("2024-05-12T15:30:00+02:00");
echo $dt->format(DateTimeInterface::ATOM) . " | " . $dt->format(DateTimeInterface::ISO8601);
