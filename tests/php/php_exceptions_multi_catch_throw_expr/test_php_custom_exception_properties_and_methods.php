<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php_custom_exception_properties_and_methods
// origin: languages/php/tests/php/test_php_exceptions_multi_catch_throw_expr.rs
// vybe-test-mode: compile

class ValidationException extends Exception {
    public function __construct(public array $errors, string $msg = "Validation failed") {
        parent::__construct($msg);
    }
}

try {
    throw new ValidationException(["email" => "Required", "age" => "Must be > 18"]);
} catch (ValidationException $e) {
    echo implode(", ", $e->errors);
}
