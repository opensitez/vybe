<?php
// vybe-test: php/php_type_juggling_coercion_strictness/test_php_is_scalar_primitive_check
// origin: languages/php/tests/php/test_php_type_juggling_coercion_strictness.rs
// vybe-test-mode: compile

echo is_scalar("str") && is_scalar(123) && is_scalar(true) && !is_scalar([]) ? "SCALAR_OK" : "FAIL";
