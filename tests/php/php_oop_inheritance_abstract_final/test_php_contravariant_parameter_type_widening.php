<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_contravariant_parameter_type_widening
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs
// vybe-test-mode: compile

class Cat {}

class Logger {
    public function log(Cat $cat): void {}
}

class UniversalLogger extends Logger {
    public function log(object $entity): void {}
}

$ul = new UniversalLogger();
$ul->log(new stdClass());
