<?php
// vybe-test: php/type_juggling/cast_bool
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump((bool) 1);
var_dump((bool) 0);
var_dump((bool) -1);
var_dump((bool) "");
var_dump((bool) "0");
var_dump((bool) "false");
var_dump((bool) []);
var_dump((bool) [0]);
var_dump((bool) null);
