<?php
// vybe-test: php/type_juggling/cast_array
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump((array) 42);
var_dump((array) "hello");
var_dump((array) null);
$obj = new stdClass(); $obj->x = 1; $obj->y = 2;
var_dump((array) $obj);
