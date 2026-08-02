<?php
// vybe-test: php/php84_datetime_create_from_timestamp/test_php84_datetime_create_from_timestamp_modify_date
// origin: languages/php/tests/php/test_php84_datetime_create_from_timestamp.rs
// vybe-test-mode: compile

$ts = 1704067200; // 2024-01-01
$dt = method_exists('DateTimeImmutable', 'createFromTimestamp')
    ? DateTimeImmutable::createFromTimestamp($ts)
    : (new DateTimeImmutable())->setTimestamp($ts);
$nextDay = $dt->modify("+1 day");
echo $nextDay->format("Y-m-d") === "2024-01-02" ? "MODIFY_TIMESTAMP_OK" : "FAIL";
