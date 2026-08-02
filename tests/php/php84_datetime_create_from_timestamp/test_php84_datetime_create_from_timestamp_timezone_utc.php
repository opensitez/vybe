<?php
// vybe-test: php/php84_datetime_create_from_timestamp/test_php84_datetime_create_from_timestamp_timezone_utc
// origin: languages/php/tests/php/test_php84_datetime_create_from_timestamp.rs
// vybe-test-mode: compile

$ts = 1700000000;
$dt = method_exists('DateTimeImmutable', 'createFromTimestamp')
    ? DateTimeImmutable::createFromTimestamp($ts)
    : (new DateTimeImmutable())->setTimestamp($ts);
echo $dt->getTimezone()->getName() === "UTC" || $dt->getTimezone()->getName() === "+00:00" ? "UTC_TIMEZONE_OK" : "FAIL";
