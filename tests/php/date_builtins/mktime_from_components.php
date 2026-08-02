<?php
// vybe-test: php/date_builtins/mktime_from_components
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

date_default_timezone_set('UTC');
$ts = mktime(0, 0, 0, 3, 15, 2024);
echo date('Y-m-d', $ts);
