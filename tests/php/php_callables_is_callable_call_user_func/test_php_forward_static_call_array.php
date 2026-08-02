<?php
// vybe-test: php/php_callables_is_callable_call_user_func/test_php_forward_static_call_array
// origin: languages/php/tests/php/test_php_callables_is_callable_call_user_func.rs
// vybe-test-mode: compile

class BaseFactory {
    public static function create(string $type) {
        return "BaseFactory:$type";
    }
}

class CustomFactory extends BaseFactory {
    public static function create(string $type) {
        return forward_static_call_array([BaseFactory::class, "create"], [$type]);
    }
}

echo CustomFactory::create("widget");
