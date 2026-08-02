<?php
// vybe-test: php/date_builtins/datetime_create_from_format_static
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$dt = DateTime::createFromFormat('d/m/Y', '25/12/2024');
echo $dt !== false ? 'created' : 'failed';
echo $dt->format('Y-m-d');
