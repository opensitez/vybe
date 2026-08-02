<?php
// vybe-test: php/php_date_time_zone_identifiers_transitions/test_php_idate_timestamp_part_extraction
// origin: languages/php/tests/php/test_php_date_time_zone_identifiers_transitions.rs
// vybe-test-mode: compile

$ts = strtotime("2024-05-12 14:30:45");
echo "Y=" . idate("Y", $ts) . " m=" . idate("m", $ts) . " d=" . idate("d", $ts);
