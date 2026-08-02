<?php
// vybe-test: php/php_date_time_zone_identifiers_transitions/test_php_date_parse_and_date_parse_from_format
// origin: languages/php/tests/php/test_php_date_time_zone_identifiers_transitions.rs
// vybe-test-mode: compile

$parsed = date_parse("2024-05-12 14:30:00");
echo "Year={$parsed['year']} Month={$parsed['month']} Day={$parsed['day']}";
