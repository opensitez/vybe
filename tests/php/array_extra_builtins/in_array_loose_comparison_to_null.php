<?php
// vybe-test: php/array_extra_builtins/in_array_loose_comparison_to_null
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = [null, false, 0, ""];
echo in_array(null, $a) ? "found" : "not";
echo in_array(null, $a, true) ? "strict-found" : "strict-not";
echo in_array(false, $a, true) ? "found" : "not";
