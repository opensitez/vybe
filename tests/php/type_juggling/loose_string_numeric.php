<?php
// vybe-test: php/type_juggling/loose_string_numeric
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump("1" == 1);       // true
var_dump("01" == 1);      // true
var_dump("1.0" == 1);     // true
var_dump("1e2" == 100);   // true
var_dump("100" == 1e2);   // true
