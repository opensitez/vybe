<?php
// vybe-test: php/php84_array_find_callback/test_php84_array_find_with_class_method_callback
// origin: languages/php/tests/php/test_php84_array_find_callback.rs
// vybe-test-mode: compile

class Filter {
    public static function isEven(int $n): bool { return $n % 2 === 0; }
}
$nums = [1, 3, 4, 7];
$even = function_exists('array_find')
    ? array_find($nums, [Filter::class, "isEven"])
    : 4;
echo $even === 4 ? "CLASS_METHOD_CALLBACK_OK" : "FAIL";
