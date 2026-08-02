<?php
// vybe-test: php/date_advanced/datetime_set_timezone_and_clone
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$dt = new DateTimeImmutable('2024-01-01 08:00:00', new DateTimeZone('UTC'));
$ny = new DateTimeZone('America/New_York');
$local = $dt->setTimezone($ny);
echo $dt->format('Y-m-d H:i') . ':' . $local->format('Y-m-d H:i');
