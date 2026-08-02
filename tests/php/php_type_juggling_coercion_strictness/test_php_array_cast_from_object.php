<?php
// vybe-test: php/php_type_juggling_coercion_strictness/test_php_array_cast_from_object
// origin: languages/php/tests/php/test_php_type_juggling_coercion_strictness.rs
// vybe-test-mode: compile

class User { public string $name = "Alice"; }
$arr = (array)(new User());
echo $arr["name"];
