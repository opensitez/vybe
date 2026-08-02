<?php
// vybe-test: php/function_builtins/func_get_args_variadic
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

function gather() {
    $args = func_get_args();
    echo count($args);
    echo implode(',', $args);
}
gather(1, 2, 3);
