<?php
// vybe-test: php/type_juggling/loose_null_equals_false
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump(null == false);   // true
var_dump(null == 0);       // true
var_dump(null == "");      // true
var_dump(null == "0");     // false
var_dump(null == []);      // true
