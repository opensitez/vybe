<?php
// vybe-test: php/filters/filter_validate_int_basic
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

var_dump(filter_var('42',   FILTER_VALIDATE_INT));
var_dump(filter_var('-10',  FILTER_VALIDATE_INT));
var_dump(filter_var('3.14', FILTER_VALIDATE_INT));
var_dump(filter_var('abc',  FILTER_VALIDATE_INT));
