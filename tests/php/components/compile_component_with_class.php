<?php
// vybe-test: php/components/compile_component_with_class
// origin: languages/php/tests/php/test_components.rs
// vybe-test-mode: compile

class Calculator {
    public function add($a, $b) { return $a + $b; }
    public function sub($a, $b) { return $a - $b; }
}
