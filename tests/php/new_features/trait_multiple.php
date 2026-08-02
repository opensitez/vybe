<?php
// vybe-test: php/new_features/trait_multiple
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

trait HasName {
    public function getName() { return $this->name; }
}
trait HasAge {
    public function getAge() { return $this->age; }
}
class User {
    use HasName;
    use HasAge;
    public $name;
    public $age;
    public function __construct($name, $age) { $this->name = $name; $this->age = $age; }
}
$u = new User("Alice", 30);
