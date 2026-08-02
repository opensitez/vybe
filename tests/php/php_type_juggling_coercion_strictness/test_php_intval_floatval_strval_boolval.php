<?php
// vybe-test: php/php_type_juggling_coercion_strictness/test_php_intval_floatval_strval_boolval
// origin: languages/php/tests/php/test_php_type_juggling_coercion_strictness.rs
// vybe-test-mode: compile

echo intval("99") + floatval("0.5") + strlen(strval(100)) + (boolval("true") ? 1 : 0);
