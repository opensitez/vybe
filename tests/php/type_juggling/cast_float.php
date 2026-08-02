<?php
// vybe-test: php/type_juggling/cast_float
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump((float) "3.14");
var_dump((float) "1e3");
var_dump((float) "abc");
var_dump((float) true);
var_dump((float) null);
