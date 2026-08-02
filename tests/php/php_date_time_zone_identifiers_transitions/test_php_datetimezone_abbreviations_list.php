<?php
// vybe-test: php/php_date_time_zone_identifiers_transitions/test_php_datetimezone_abbreviations_list
// origin: languages/php/tests/php/test_php_date_time_zone_identifiers_transitions.rs
// vybe-test-mode: compile

$abbrevs = DateTimeZone::listAbbreviations();
echo isset($abbrevs["est"]) && isset($abbrevs["pst"]) ? "ABBREVS_OK" : "FAIL";
