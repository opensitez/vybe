<?php
// vybe-test: php/functional_style/array_filter_arrow_function
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

$nums = [1, 2, 3, 4, 5, 6, 7, 8];
$evens = array_filter($nums, fn($n) => $n % 2 === 0);
echo implode(',', array_values($evens));
