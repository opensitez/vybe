<?php
// vybe-test: php/date_builtins/date_modify_object
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$dt = date_create('2024-01-01');
date_modify($dt, '+3 months');
echo date_format($dt, 'Y-m-d');
