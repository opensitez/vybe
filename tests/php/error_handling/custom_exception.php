<?php
// vybe-test: php/error_handling/custom_exception
// origin: languages/php/tests/php/test_error_handling.rs
// vybe-test-mode: compile

class AppException extends Exception {
    public $code;
    public function __construct($message, $code = 0) {
        $this->code = $code;
    }
}
try {
    throw new AppException('not found', 404);
} catch (AppException $e) {
    echo $e->code;
}
