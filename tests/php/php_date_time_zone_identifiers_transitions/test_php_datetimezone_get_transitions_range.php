<?php
// vybe-test: php/php_date_time_zone_identifiers_transitions/test_php_datetimezone_get_transitions_range
// origin: languages/php/tests/php/test_php_date_time_zone_identifiers_transitions.rs
// vybe-test-mode: compile

$tz = new DateTimeZone("Europe/Berlin");
$transitions = $tz->getTransitions(
    strtotime("2024-01-01"),
    strtotime("2024-12-31")
);
echo "Transitions count: " . count($transitions);
