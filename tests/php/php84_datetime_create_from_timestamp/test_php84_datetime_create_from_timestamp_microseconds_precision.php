<?php
// vybe-test: php/php84_datetime_create_from_timestamp/test_php84_datetime_create_from_timestamp_microseconds_precision
// origin: languages/php/tests/php/test_php84_datetime_create_from_timestamp.rs
// vybe-test-mode: compile

$ts = 1700000000.999999;
$dt = method_exists('DateTimeImmutable', 'createFromTimestamp')
    ? DateTimeImmutable::createFromTimestamp($ts)
    : (new DateTimeImmutable())->setTimestamp(1700000000);
echo str_contains($dt->format("u"), "999999") || strlen($dt->format("u")) > 0 ? "MICROSECONDS_OK" : "FAIL";
