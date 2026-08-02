<?php
// vybe-test: php/php_date_time_zone_identifiers_transitions/test_php_checkdate_gregorian_validation
// origin: languages/php/tests/php/test_php_date_time_zone_identifiers_transitions.rs
// vybe-test-mode: compile

echo checkdate(2, 29, 2024) ? "LEAP_FEB29_OK" : "FAIL"; // 2024 is a leap year
echo !checkdate(2, 29, 2023) ? "FEB29_INVALID" : "FAIL";
