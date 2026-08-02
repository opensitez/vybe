<?php
// vybe-test: php/filters/filter_validate_float
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

var_dump(filter_var('3.14',   FILTER_VALIDATE_FLOAT));
var_dump(filter_var('1e5',    FILTER_VALIDATE_FLOAT));
var_dump(filter_var('-0.001', FILTER_VALIDATE_FLOAT));
var_dump(filter_var('abc',    FILTER_VALIDATE_FLOAT));
