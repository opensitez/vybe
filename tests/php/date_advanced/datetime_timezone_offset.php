<?php
// vybe-test: php/date_advanced/datetime_timezone_offset
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$tz  = new DateTimeZone('UTC');
$dt  = new DateTime('now', $tz);
$off = $tz->getOffset($dt);
echo $off;  // 0 for UTC
