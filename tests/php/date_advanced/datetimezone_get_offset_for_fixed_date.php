<?php
// vybe-test: php/date_advanced/datetimezone_get_offset_for_fixed_date
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$tz = new DateTimeZone('UTC');
$dt = new DateTimeImmutable('2024-03-01 00:00:00', $tz);
echo $tz->getOffset($dt);
