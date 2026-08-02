<?php
// vybe-test: php/type_juggling/coercion_bool_arithmetic
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump(true  + true);   // int(2)
var_dump(false + 1);      // int(1)
var_dump(true  + 0.5);    // float(1.5)
var_dump(true  * 10);     // int(10)
