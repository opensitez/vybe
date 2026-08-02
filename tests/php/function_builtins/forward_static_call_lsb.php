<?php
// vybe-test: php/function_builtins/forward_static_call_lsb
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

class Base {
    static function create() {
        return forward_static_call(['static', 'build']);
    }
    static function build() {
        return 'base';
    }
}
class Child extends Base {
    static function build() {
        return 'child';
    }
}
echo Child::create();
