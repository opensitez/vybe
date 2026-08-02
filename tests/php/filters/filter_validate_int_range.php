<?php
// vybe-test: php/filters/filter_validate_int_range
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$options = ['options' => ['min_range' => 1, 'max_range' => 100]];
var_dump(filter_var('50',  FILTER_VALIDATE_INT, $options));
var_dump(filter_var('0',   FILTER_VALIDATE_INT, $options));
var_dump(filter_var('101', FILTER_VALIDATE_INT, $options));
