<?php
// vybe-test: php/magic_constants/magic_method_basic
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

class Calculator {
    public function add(): string { return __METHOD__; }
}
echo (new Calculator())->add();
