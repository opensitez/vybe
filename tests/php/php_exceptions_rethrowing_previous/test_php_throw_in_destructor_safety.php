<?php
// vybe-test: php/php_exceptions_rethrowing_previous/test_php_throw_in_destructor_safety
// origin: languages/php/tests/php/test_php_exceptions_rethrowing_previous.rs
// vybe-test-mode: compile

class DangerousDestructor {
    public function __destruct() {
        // Exceptions thrown from destructors must be handled or will trigger fatal error
    }
}
$d = new DangerousDestructor();
