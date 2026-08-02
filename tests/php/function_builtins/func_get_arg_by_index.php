<?php
// vybe-test: php/function_builtins/func_get_arg_by_index
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

function get_second() {
    return func_get_arg(1);
}
echo get_second('first', 'second', 'third');
