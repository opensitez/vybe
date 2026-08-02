<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php_throw_expression_in_match_arm
// origin: languages/php/tests/php/test_php_exceptions_multi_catch_throw_expr.rs
// vybe-test-mode: compile

$action = "invalid";
$res = match ($action) {
    "run" => "running",
    "stop" => "stopped",
    default => throw new DomainException("Invalid action: $action"),
};
