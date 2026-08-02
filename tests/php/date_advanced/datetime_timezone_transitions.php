<?php
// vybe-test: php/date_advanced/datetime_timezone_transitions
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$tz = new DateTimeZone('America/New_York');
$transitions = $tz->getTransitions(mktime(0,0,0,1,1,2024), mktime(0,0,0,12,31,2024));
echo count($transitions) > 0 ? 'has transitions' : 'no transitions';
