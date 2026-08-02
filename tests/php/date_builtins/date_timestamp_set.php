<?php
// vybe-test: php/date_builtins/date_timestamp_set
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

date_default_timezone_set('UTC');
$dt = date_create('now');
$ts = mktime(0, 0, 0, 6, 15, 2024);
date_timestamp_set($dt, $ts);
echo date_format($dt, 'Y-m-d');
