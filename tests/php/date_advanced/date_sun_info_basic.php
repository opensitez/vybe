<?php
// vybe-test: php/date_advanced/date_sun_info_basic
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$info = date_sun_info(mktime(0,0,0,6,21,2024), 51.5, -0.12); // London midsummer
echo isset($info['sunrise'])  ? 'has sunrise' : 'no sunrise';
echo isset($info['sunset'])   ? ':has sunset' : ':no sunset';
echo isset($info['transit'])  ? ':has transit' : ':no transit';
