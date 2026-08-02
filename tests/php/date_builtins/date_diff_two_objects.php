<?php
// vybe-test: php/date_builtins/date_diff_two_objects
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$d1 = date_create('2024-01-01');
$d2 = date_create('2024-06-15');
$diff = date_diff($d1, $d2);
echo $diff->days > 0 ? 'positive days' : 'zero or negative';
echo $diff->m > 0 ? ':has months' : ':no months';
