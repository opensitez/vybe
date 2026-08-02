<?php
// vybe-test: php/magic_constants/magic_class_basic
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

class MyClass {
    public function getClass(): string { return __CLASS__; }
}
echo (new MyClass())->getClass();
