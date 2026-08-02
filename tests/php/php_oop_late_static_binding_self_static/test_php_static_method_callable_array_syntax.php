<?php
// vybe-test: php/php_oop_late_static_binding_self_static/test_php_static_method_callable_array_syntax
// origin: languages/php/tests/php/test_php_oop_late_static_binding_self_static.rs
// vybe-test-mode: compile

class Dispatcher {
    public static function handle(string $event) { return "Handled: $event"; }
}

$callable = [Dispatcher::class, "handle"];
echo call_user_func($callable, "login");
