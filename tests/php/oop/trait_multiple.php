<?php
// vybe-test: php/oop/trait_multiple
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

trait HasName { public function getName() { return $this->name; } }
trait HasAge { public function getAge() { return $this->age; } }
class User {
    use HasName;
    use HasAge;
    public $name; public $age;
    public function __construct($n, $a) { $this->name = $n; $this->age = $a; }
}
$u = new User('Bob', 25);
echo $u->getName();
