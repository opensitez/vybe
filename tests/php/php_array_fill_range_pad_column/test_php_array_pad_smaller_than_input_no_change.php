<?php
// vybe-test: php/php_array_fill_range_pad_column/test_php_array_pad_smaller_than_input_no_change
// origin: languages/php/tests/php/test_php_array_fill_range_pad_column.rs
// vybe-test-mode: compile

$a = [1, 2, 3];
$padded = array_pad($a, 2, "x");
echo count($padded) === 3 ? "NO_PAD_OK" : "FAIL";
