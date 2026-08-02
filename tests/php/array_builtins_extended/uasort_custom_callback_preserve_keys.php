<?php
// vybe-test: php/array_builtins_extended/uasort_custom_callback_preserve_keys
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["p" => 30, "q" => 10, "r" => 20];
uasort($a, function($x, $y) { return $x - $y; });
echo implode(",", array_keys($a));
echo implode(",", $a);
