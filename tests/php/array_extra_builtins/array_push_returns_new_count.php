<?php
// vybe-test: php/array_extra_builtins/array_push_returns_new_count
// origin: languages/php/tests/php/test_array_extra_builtins.rs
// vybe-test-mode: compile

$a = [1, 2, 3];
$count = array_push($a, 4, 5, 6);
echo $count;
echo count($a);
echo $a[5];
