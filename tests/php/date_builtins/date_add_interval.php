<?php
// vybe-test: php/date_builtins/date_add_interval
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$dt = date_create('2024-01-01');
$interval = new DateInterval('P2M');
date_add($dt, $interval);
echo date_format($dt, 'Y-m');
