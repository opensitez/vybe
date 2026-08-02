<?php
// vybe-test: php/type_juggling/cast_int
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump((int) "42");
var_dump((int) "42abc");
var_dump((int) "abc");
var_dump((int) 3.9);
var_dump((int) true);
var_dump((int) null);
var_dump((int) false);
