<?php
// vybe-test: php/date_builtins/date_sub_interval
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$dt = date_create('2024-06-15');
$interval = new DateInterval('P15D');
date_sub($dt, $interval);
echo date_format($dt, 'Y-m-d');
