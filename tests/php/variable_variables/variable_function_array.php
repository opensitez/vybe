<?php
// vybe-test: php/variable_variables/variable_function_array
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

function square(int $n): int { return $n * $n; }
function cube(int $n): int   { return $n * $n * $n; }
$ops = ['square', 'cube'];
foreach ($ops as $fn) { echo $fn(3) . ' '; }
