<?php
// vybe-test: php/type_juggling/loose_zero_string_comparison
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump(0 == "a");    // true in PHP 7, false in PHP 8
var_dump(0 == "");     // true in PHP 7, false in PHP 8
var_dump(0 == "0");    // true
var_dump(0 == false);  // true
var_dump(0 == null);   // true
