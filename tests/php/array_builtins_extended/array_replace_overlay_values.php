<?php
// vybe-test: php/array_builtins_extended/array_replace_overlay_values
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$base    = ["a" => 1, "b" => 2, "c" => 3];
$overlay = ["b" => 20, "d" => 40];
$result  = array_replace($base, $overlay);
echo $result["a"];
echo $result["b"];
echo $result["d"];
echo array_key_exists("c", $result) ? "yes" : "no";
