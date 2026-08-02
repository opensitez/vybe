<?php
// vybe-test: php/date_builtins/date_timestamp_get
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

date_default_timezone_set('UTC');
$dt = date_create('2024-01-01 00:00:00');
$ts = date_timestamp_get($dt);
echo is_int($ts) ? 'integer' : 'not integer';
echo date('Y', $ts);
