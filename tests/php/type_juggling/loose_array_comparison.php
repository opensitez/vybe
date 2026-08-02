<?php
// vybe-test: php/type_juggling/loose_array_comparison
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump([] == false);   // true
var_dump([] == null);    // true
var_dump([] == 0);       // false
var_dump([0] == [false]); // true
var_dump(['a' => 1] == ['a' => true]); // true
