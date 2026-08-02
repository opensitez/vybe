<?php
// vybe-test: php/date_advanced/datetime_convert_timezone
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$utc = new DateTimeImmutable('2024-06-15 12:00:00', new DateTimeZone('UTC'));
$ny  = $utc->setTimezone(new DateTimeZone('America/New_York'));
$tok = $utc->setTimezone(new DateTimeZone('Asia/Tokyo'));
// NY is UTC-4 in summer, Tokyo is UTC+9
echo $ny->format('H')  . ':' . $tok->format('H');
