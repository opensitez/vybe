<?php
// vybe-test: php/class_inspection/trait_exists_defined_trait
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

trait Greetable { public function greet() { return 'hi'; } }
echo trait_exists('Greetable') ? 'yes' : 'no';
echo trait_exists('Missing') ? 'yes' : 'no';
