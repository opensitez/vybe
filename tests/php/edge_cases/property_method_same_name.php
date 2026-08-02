<?php
// vybe-test: php/edge_cases/property_method_same_name
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

class A { public $name = 'prop'; public function name() { return 'method'; } } $a = new A(); echo $a->name; echo $a->name();
