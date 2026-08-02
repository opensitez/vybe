<?php
// vybe-test: php/oop/trait_basic
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

trait Greetable {
    public function greet() { return 'Hello, ' . $this->name; }
}
class Person { use Greetable; public $name; public function __construct($n) { $this->name = $n; } }
$p = new Person('Alice');
echo $p->greet();
