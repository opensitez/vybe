<?php
// vybe-test: php/exception_types/custom_exception_hierarchy
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

class AppException extends RuntimeException {}
class NetworkException extends AppException {
    public function __construct(string $host) {
        parent::__construct('unreachable: ' . $host, 503);
    }
}
throw new NetworkException('example.com');
