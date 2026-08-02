<?php
// vybe-test: php/type_juggling/coercion_null_arithmetic
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump(null + 1);     // int(1)
var_dump(null + 1.5);   // float(1.5)
var_dump(null . "str"); // string "str"
