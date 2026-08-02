<?php
// vybe-test: php/array_extra_builtins/array_map_with_static_callables
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$vals = [" 1", " 2", " 3"];
$trimmed = array_map("trim", $vals);
echo implode(",", $trimmed);
