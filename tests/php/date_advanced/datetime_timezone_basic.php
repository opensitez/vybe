<?php
// vybe-test: php/date_advanced/datetime_timezone_basic
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$tz = new DateTimeZone('America/New_York');
echo $tz->getName();
