<?php
// vybe-test: php/php_type_juggling_coercion_strictness/test_php_gettype_builtin_returns
// origin: languages/php/tests/php/test_php_type_juggling_coercion_strictness.rs
// vybe-test-mode: compile

echo gettype(1) . " " . gettype(1.0) . " " . gettype("a") . " " . gettype([]) . " " . gettype(null);
