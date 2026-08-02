<?php
// vybe-test: php/variable_functions/list_from_function_return
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

function minMax(array $arr): array {
    return [min($arr), max($arr)];
}
[$lo, $hi] = minMax([3, 1, 4, 1, 5, 9]);
echo $lo;
echo $hi;
