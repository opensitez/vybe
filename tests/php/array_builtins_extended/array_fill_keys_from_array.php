<?php
// vybe-test: php/array_builtins_extended/array_fill_keys_from_array
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$keys = ["alpha", "beta", "gamma"];
$a = array_fill_keys($keys, null);
echo count($a);
echo array_key_exists("beta", $a) ? "yes" : "no";
echo array_key_exists("delta", $a) ? "yes" : "no";
