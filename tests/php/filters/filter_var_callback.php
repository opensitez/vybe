<?php
// vybe-test: php/filters/filter_var_callback
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$result = filter_var('hello world', FILTER_CALLBACK, ['options' => 'strtoupper']);
echo $result;
