<?php
// vybe-test: php/php_date_time_zone_identifiers_transitions/test_php_date_sun_info_sunrise_sunset
// origin: languages/php/tests/php/test_php_date_time_zone_identifiers_transitions.rs
// vybe-test-mode: compile

$sunInfo = date_sun_info(strtotime("2024-06-21"), 51.5, -0.12); // London summer solstice
echo "Sunrise=" . date("H:i", $sunInfo["sunrise"]) . " Sunset=" . date("H:i", $sunInfo["sunset"]);
