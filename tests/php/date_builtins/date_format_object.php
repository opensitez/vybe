<?php
// vybe-test: php/date_builtins/date_format_object
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$dt = date_create('2024-12-25');
echo date_format($dt, 'Y-m-d');
echo date_format($dt, 'l');
