<?php
// vybe-test: php/variable_functions/empty_on_falsy_values
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$checks = [0, '', [], null, false, '0', 1, 'x', [1], true];
foreach ($checks as $v) {
    echo empty($v) ? '1' : '0';
}
