<?php
// vybe-test: php/array_extra_builtins/array_fill_keys_and_counting
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$map = array_fill_keys([1, 2, 3], "x");
echo $map[1];
echo $map[2];
echo $map[3];
echo count($map);
