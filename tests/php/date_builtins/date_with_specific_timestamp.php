<?php
// vybe-test: php/date_builtins/date_with_specific_timestamp
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$ts = 1704067200; // 2024-01-01 00:00:00 UTC
date_default_timezone_set('UTC');
echo date('Y', $ts);
echo date('m', $ts);
echo date('d', $ts);
