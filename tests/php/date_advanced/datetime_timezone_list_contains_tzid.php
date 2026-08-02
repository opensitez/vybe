<?php
// vybe-test: php/date_advanced/datetime_timezone_list_contains_tzid
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$zones = DateTimeZone::listIdentifiers();
echo in_array('UTC', $zones) ? 'has-utc' : 'missing-utc';
echo '|' . (DateTimeZone::listIdentifiers(DateTimeZone::AMERICA)[0] !== '' ? 'region' : 'no-region');
echo '|' . (new DateTimeZone('UTC'))->getName();
