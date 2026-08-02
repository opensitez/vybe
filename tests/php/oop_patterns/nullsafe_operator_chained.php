<?php
// vybe-test: php/oop_patterns/nullsafe_operator_chained
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Street {
    public function __construct(public string $name) {}
}
class Address {
    public function __construct(public ?Street $street = null) {}
    public function getStreet(): ?Street { return $this->street; }
}
class User {
    public function __construct(public ?Address $address = null) {}
    public function getAddress(): ?Address { return $this->address; }
}
$userWithAddress    = new User(new Address(new Street('Main St')));
$userWithoutAddress = new User(null);
echo $userWithAddress?->getAddress()?->getStreet()?->name ?? 'none';
echo $userWithoutAddress?->getAddress()?->getStreet()?->name ?? 'none';
