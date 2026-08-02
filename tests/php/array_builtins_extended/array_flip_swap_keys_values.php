<?php
// vybe-test: php/array_builtins_extended/array_flip_swap_keys_values
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["one" => 1, "two" => 2, "three" => 3];
$flipped = array_flip($a);
echo $flipped[1];
echo $flipped[2];
echo array_key_exists("one", $flipped) ? "bad" : "ok";
