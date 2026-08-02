<?php
// vybe-test: php/php_datetime_immutable_interval_period/test_php_dateinterval_create_from_datestring
// origin: languages/php/tests/php/test_php_datetime_immutable_interval_period.rs
// vybe-test-mode: compile

$interval = DateInterval::createFromDateString("2 days 3 hours");
$dt = (new DateTimeImmutable("2024-01-01 00:00:00"))->add($interval);
echo $dt->format("Y-m-d H:i:s");
