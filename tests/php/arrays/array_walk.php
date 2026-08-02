<?php
// vybe-test: php/arrays/array_walk
// origin: languages/php/tests/php/test_arrays.rs
// vybe-test-mode: compile

$a = [1,2,3]; array_walk($a, fn($v, $k) => $v);
