<?php
// vybe-test: php/new_features/parent_method_call
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

class Base {
    public function hello() { return "Hello from Base"; }
}
class Child extends Base {
    public function hello() { return parent::hello() . " and Child"; }
}
$c = new Child();
