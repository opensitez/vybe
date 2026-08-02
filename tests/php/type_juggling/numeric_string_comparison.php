<?php
// vybe-test: php/type_juggling/numeric_string_comparison
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

// When both strings are numeric, PHP compares numerically
var_dump("1" < "10");    // true (numeric)
var_dump("abc" < "abd"); // true (string)
var_dump("2" > "10");    // false (numeric: 2 < 10)
