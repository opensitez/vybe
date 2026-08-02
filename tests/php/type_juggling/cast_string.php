<?php
// vybe-test: php/type_juggling/cast_string
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

var_dump((string) 42);
var_dump((string) 3.14);
var_dump((string) true);
var_dump((string) false);
var_dump((string) null);
