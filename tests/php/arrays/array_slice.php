<?php
// vybe-test: php/arrays/array_slice
// origin: languages/php/tests/php/test_arrays.rs
// vybe-test-mode: compile

$x = array_slice([1, 2, 3, 4, 5], 1, 3); echo count($x);
