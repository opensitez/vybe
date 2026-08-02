<?php
// vybe-test: php/arrays/array_reduce
// origin: languages/php/tests/php/test_arrays.rs
// vybe-test-mode: compile

$sum = array_reduce([1,2,3], fn($carry, $item) => $carry + $item, 0);
