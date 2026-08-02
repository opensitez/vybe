<?php
// vybe-test: php/php_type_juggling_coercion_strictness/test_php_settype_in_place_conversion
// origin: languages/php/tests/php/test_php_type_juggling_coercion_strictness.rs
// vybe-test-mode: compile

$foo = "5bar";
settype($foo, "integer");
echo $foo === 5 ? "SETTYPE_OK" : "FAIL";
