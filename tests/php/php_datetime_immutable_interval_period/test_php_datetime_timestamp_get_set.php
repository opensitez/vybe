<?php
// vybe-test: php/php_datetime_immutable_interval_period/test_php_datetime_timestamp_get_set
// origin: languages/php/tests/php/test_php_datetime_immutable_interval_period.rs
// vybe-test-mode: compile

$dt = new DateTimeImmutable();
$ts = $dt->getTimestamp();
$dt2 = $dt->setTimestamp($ts + 3600);
echo ($dt2->getTimestamp() - $ts);
