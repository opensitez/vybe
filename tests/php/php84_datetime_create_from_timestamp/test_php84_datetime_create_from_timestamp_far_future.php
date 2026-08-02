<?php
// vybe-test: php/php84_datetime_create_from_timestamp/test_php84_datetime_create_from_timestamp_far_future
// origin: languages/php/tests/php/test_php84_datetime_create_from_timestamp.rs
// vybe-test-mode: compile

$ts = 2524608000; // 2050-01-01
$dt = method_exists('DateTimeImmutable', 'createFromTimestamp')
    ? DateTimeImmutable::createFromTimestamp($ts)
    : (new DateTimeImmutable())->setTimestamp($ts);
echo $dt->format("Y") === "2050" ? "FAR_FUTURE_OK" : "FAIL";
