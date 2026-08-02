<?php
// vybe-test: php/php84_datetime_create_from_timestamp/test_php84_datetime_create_from_timestamp_diff_calculation
// origin: languages/php/tests/php/test_php84_datetime_create_from_timestamp.rs
// vybe-test-mode: compile

$ts1 = 1704067200; // 2024-01-01
$ts2 = 1704153600; // 2024-01-02
$dt1 = method_exists('DateTimeImmutable', 'createFromTimestamp') ? DateTimeImmutable::createFromTimestamp($ts1) : (new DateTimeImmutable())->setTimestamp($ts1);
$dt2 = method_exists('DateTimeImmutable', 'createFromTimestamp') ? DateTimeImmutable::createFromTimestamp($ts2) : (new DateTimeImmutable())->setTimestamp($ts2);
$diff = $dt1->diff($dt2);
echo $diff->days === 1 ? "TIMESTAMP_DIFF_1DAY_OK" : "FAIL";
