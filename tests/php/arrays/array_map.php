<?php
// vybe-test: php/arrays/array_map
// origin: languages/php/tests/php/test_arrays.rs
// vybe-test-mode: compile

$x = array_map(fn($n) => $n * 2, [1, 2, 3]);
