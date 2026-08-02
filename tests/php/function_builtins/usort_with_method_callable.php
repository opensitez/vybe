<?php
// vybe-test: php/function_builtins/usort_with_method_callable
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

class Comparator {
    public function compare($a, $b): int {
        return $a <=> $b;
    }
}
$cmp = new Comparator();
$arr = [3, 1, 4, 1, 5, 9, 2, 6];
usort($arr, [$cmp, 'compare']);
echo implode(',', $arr);
