<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php_throw_expression_in_arrow_function
// origin: languages/php/tests/php/test_php_exceptions_multi_catch_throw_expr.rs
// vybe-test-mode: compile

$validate = fn($val) => $val > 0 ? $val : throw new InvalidArgumentException("Must be positive");
echo $validate(10);
