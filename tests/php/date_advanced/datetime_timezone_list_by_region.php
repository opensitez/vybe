<?php
// vybe-test: php/date_advanced/datetime_timezone_list_by_region
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$zones = DateTimeZone::listIdentifiers(DateTimeZone::EUROPE);
echo count($zones) > 0 ? 'has zones' : 'empty';
echo in_array('Europe/London', $zones) ? ':has London' : ':no London';
