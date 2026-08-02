<?php
// vybe-test: php/type_juggling/is_numeric_check
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump(is_numeric(42));
var_dump(is_numeric(3.14));
var_dump(is_numeric("42"));
var_dump(is_numeric("3.14"));
var_dump(is_numeric("1e5"));
var_dump(is_numeric("42abc"));
var_dump(is_numeric("abc"));
var_dump(is_numeric(""));
var_dump(is_numeric(null));
