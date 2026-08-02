<?php
// vybe-test: php/php_array_fill_range_pad_column/test_php_array_fill_zero_num_returns_empty
// origin: languages/php/tests/php/test_php_array_fill_range_pad_column.rs
// vybe-test-mode: compile

$res = array_fill(0, 0, "val");
echo count($res) === 0 ? "ZERO_NUM_OK" : "FAIL";
