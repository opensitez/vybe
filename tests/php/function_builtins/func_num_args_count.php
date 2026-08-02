<?php
// vybe-test: php/function_builtins/func_num_args_count
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

function count_args() {
    return func_num_args();
}
echo count_args(10, 20, 30);
echo count_args();
