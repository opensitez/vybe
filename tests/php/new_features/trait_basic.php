<?php
// vybe-test: php/new_features/trait_basic
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

trait Greetable {
    public function greet() { return "Hello from " . $this->name; }
}
class Person {
    use Greetable;
    public $name;
    public function __construct($name) { $this->name = $name; }
}
$p = new Person("John");
echo $p->greet();
