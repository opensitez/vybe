<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_abstract_property_visibility_preserved_runtime
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs
// vybe-test-mode: compile

abstract class UserEntity {
    public function __construct(protected string $name) {}
    protected function getName(): string { return $this->name; }
}
class CustomerEntity extends UserEntity {
    public function label(): string { return $this->getName(); }
}
echo (new CustomerEntity('acme'))->label();
