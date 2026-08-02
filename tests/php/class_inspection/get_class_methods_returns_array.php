<?php
// vybe-test: php/class_inspection/get_class_methods_returns_array
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class Calc {
    public function add($a, $b) { return $a + $b; }
    public function sub($a, $b) { return $a - $b; }
}
$methods = get_class_methods('Calc');
echo in_array('add', $methods) ? 'yes' : 'no';
echo in_array('sub', $methods) ? 'yes' : 'no';
