<?php
// vybe-test: php/date_builtins/date_sunrise_sunset_compat
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$sunrise = date_sunrise(strtotime('2024-06-21'), SUNFUNCS_RET_STRING, 40.7128, -74.0060, 90, -5);
$sunset = date_sunset(strtotime('2024-06-21'), SUNFUNCS_RET_STRING, 40.7128, -74.0060, 90, -5);
echo is_string($sunrise) ? 'sr' : 'ns';
echo ':' . (is_string($sunset) ? 'ss' : 'ns2');
