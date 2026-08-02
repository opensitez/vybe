<?php
// vybe-test: php/class_inspection/method_exists_class_and_object
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class Greeter {
    public function hello(): string { return 'hi'; }
}
$g = new Greeter();
echo method_exists($g, 'hello') ? 'yes' : 'no';
echo method_exists($g, 'goodbye') ? 'yes' : 'no';
echo method_exists('Greeter', 'hello') ? 'yes' : 'no';
