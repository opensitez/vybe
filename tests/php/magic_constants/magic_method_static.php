<?php
// vybe-test: php/magic_constants/magic_method_static
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

class Factory {
    public static function create(): string { return __METHOD__; }
}
echo Factory::create();
