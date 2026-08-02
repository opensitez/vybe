<?php
// vybe-test: php/oop_patterns/clone_with_deep_copy
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Address {
    public function __construct(public string $city) {}
}
class Person {
    public function __construct(public string $name, public Address $address) {}
    public function __clone() {
        $this->address = clone $this->address;
    }
}
$original = new Person('Alice', new Address('Paris'));
$copy = clone $original;
$copy->address->city = 'London';
echo $original->address->city;
echo $copy->address->city;
