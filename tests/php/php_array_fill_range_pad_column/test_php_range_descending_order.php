<?php
// vybe-test: php/php_array_fill_range_pad_column/test_php_range_descending_order
// origin: languages/php/tests/php/test_php_array_fill_range_pad_column.rs
// vybe-test-mode: compile

$desc = range(10, 1);
echo count($desc) === 10 && $desc[0] === 10 ? "DESC_RANGE_OK" : "FAIL";
