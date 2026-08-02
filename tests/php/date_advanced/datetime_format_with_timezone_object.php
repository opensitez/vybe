<?php
// vybe-test: php/date_advanced/datetime_format_with_timezone_object
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$tz = new DateTimeZone('America/Chicago');
$dt = new DateTimeImmutable('2024-01-01 12:00:00', $tz);
echo $dt->format('c');
