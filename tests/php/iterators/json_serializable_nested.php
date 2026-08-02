<?php
// vybe-test: php/iterators/json_serializable_nested
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class Address implements JsonSerializable {
    public function __construct(public string $city, public string $country) {}
    public function jsonSerialize(): array { return ['city' => $this->city, 'country' => $this->country]; }
}
class Person implements JsonSerializable {
    public function __construct(public string $name, public Address $address) {}
    public function jsonSerialize(): array {
        return ['name' => $this->name, 'address' => $this->address];
    }
}
$p = new Person('Alice', new Address('Paris', 'FR'));
echo json_encode($p);
