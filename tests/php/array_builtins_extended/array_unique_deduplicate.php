<?php
// vybe-test: php/array_builtins_extended/array_unique_deduplicate
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["a", "b", "a", "c", "b", "d", "d"];
$u = array_unique($a);
echo count($u);
echo implode(",", $u);
