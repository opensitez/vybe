<?php
// vybe-test: php/date_builtins/date_create_alias
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$dt = date_create('2024-06-15');
echo $dt !== false ? 'created' : 'failed';
echo date_format($dt, 'Y');
