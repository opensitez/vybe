<?php
// vybe-test: php/magic_constants/magic_class_in_static
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

class Foo {
    public static function name(): string { return __CLASS__; }
}
echo Foo::name();
