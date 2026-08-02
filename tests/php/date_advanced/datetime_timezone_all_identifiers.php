<?php
// vybe-test: php/date_advanced/datetime_timezone_all_identifiers
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$all = DateTimeZone::listIdentifiers();
echo count($all) > 400 ? 'many zones' : 'few zones';
