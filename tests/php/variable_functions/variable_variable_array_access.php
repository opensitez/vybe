<?php
// vybe-test: php/variable_functions/variable_variable_array_access
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$fields = ['x', 'y', 'z'];
$x = 10;
$y = 20;
$z = 30;
$sum = 0;
foreach ($fields as $name) {
    $sum += $$name;
}
echo $sum;
