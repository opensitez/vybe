<?php
// vybe-test: php/type_juggling/loose_bool_comparison
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump(true  == 1);      // true
var_dump(true  == "1");    // true
var_dump(true  == "any");  // true
var_dump(false == 0);      // true
var_dump(false == "");     // true
var_dump(false == "0");    // true
var_dump(false == []);     // true
var_dump(false == null);   // true
