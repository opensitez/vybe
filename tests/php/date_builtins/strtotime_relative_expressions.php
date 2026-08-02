<?php
// vybe-test: php/date_builtins/strtotime_relative_expressions
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

date_default_timezone_set('UTC');
$base = strtotime('2024-01-01');
$next_week = strtotime('+1 week', $base);
echo date('d', $next_week);
$next_month = strtotime('+1 month', $base);
echo date('m', $next_month);
