<?php
// vybe-test: php/php_oop_late_static_binding_self_static/test_php_forward_static_call_forwarding
// origin: languages/php/tests/php/test_php_oop_late_static_binding_self_static.rs
// vybe-test-mode: compile

class A {
    public static function foo() {
        echo get_called_class();
    }
}
class B extends A {
    public static function foo() {
        forward_static_call(['A', 'foo']);
    }
}

B::foo();
