<?php
// vybe-test: php/array_builtins_extended/array_fill_with_start_index
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = array_fill(5, 4, "x");
echo count($a);
echo $a[5];
echo $a[8];
echo array_key_exists(4, $a) ? "bad" : "ok";
