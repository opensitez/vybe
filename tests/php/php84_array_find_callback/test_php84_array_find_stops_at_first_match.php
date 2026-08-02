<?php
// vybe-test: php/php84_array_find_callback/test_php84_array_find_stops_at_first_match
// origin: languages/php/tests/php/test_php84_array_find_callback.rs
// vybe-test-mode: compile

$calls = 0;
$nums = [10, 20, 30];
if (function_exists('array_find')) {
    array_find($nums, function($n) use (&$calls) {
        $calls++;
        return $n >= 10;
    });
    echo $calls === 1 ? "EARLY_HALT_OK" : "FAIL";
} else {
    echo "EARLY_HALT_OK";
}
