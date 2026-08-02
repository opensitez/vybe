<?php
// vybe-test: php/php_oop_late_static_binding_self_static/test_php_static_return_type_hint_php80
// origin: languages/php/tests/php/test_php_oop_late_static_binding_self_static.rs
// vybe-test-mode: compile

class Chainable {
    public function setOption(): static {
        return $this;
    }
}

class SubChainable extends Chainable {}

$sc = new SubChainable();
echo get_class($sc->setOption());
