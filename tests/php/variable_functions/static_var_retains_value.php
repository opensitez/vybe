<?php
// vybe-test: php/variable_functions/static_var_retains_value
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

function counter(): int {
    static $count = 0;
    $count++;
    return $count;
}
echo counter();
echo counter();
echo counter();
