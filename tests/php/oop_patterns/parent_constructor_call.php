<?php
// vybe-test: php/oop_patterns/parent_constructor_call
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Vehicle {
    public function __construct(
        protected string $make,
        protected int    $year
    ) {}
    public function info(): string { return "{$this->year} {$this->make}"; }
}
class Car extends Vehicle {
    public function __construct(
        string $make,
        int    $year,
        private int $doors
    ) {
        parent::__construct($make, $year);
    }
    public function describe(): string { return $this->info() . " ({$this->doors} doors)"; }
}
$car = new Car('Toyota', 2023, 4);
echo $car->describe();
