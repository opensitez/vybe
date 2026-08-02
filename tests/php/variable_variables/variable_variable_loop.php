<?php
// vybe-test: php/variable_variables/variable_variable_loop
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

$vars = ['a', 'b', 'c'];
foreach ($vars as $i => $name) {
    $$name = $i * 10;
}
echo $a . ',' . $b . ',' . $c;
