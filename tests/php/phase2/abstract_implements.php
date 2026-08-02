<?php
// vybe-test: php/phase2/abstract_implements
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

abstract class Vehicle {
    abstract public function wheels(): int;
    public function describe() { return 'Vehicle with ' . $this->wheels() . ' wheels'; }
}
class Car extends Vehicle {
    public function wheels(): int { return 4; }
}
$car = new Car();
