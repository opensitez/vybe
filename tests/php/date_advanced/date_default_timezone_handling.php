<?php
// vybe-test: php/date_advanced/date_default_timezone_handling
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$original = date_default_timezone_get();
date_default_timezone_set('Europe/London');
$tz = date_default_timezone_get();
echo $tz === 'Europe/London' ? 'ok' : 'no';
date_default_timezone_set($original);
