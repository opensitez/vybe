<?php
// vybe-test: php/function_builtins/usort_with_static_method_callable
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

class Sorter {
    public static function descending($a, $b): int {
        return $b <=> $a;
    }
}
$arr = [3, 1, 4, 1, 5, 9, 2, 6];
usort($arr, ['Sorter', 'descending']);
echo implode(',', $arr);
