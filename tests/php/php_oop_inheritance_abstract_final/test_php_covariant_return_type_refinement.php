<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_covariant_return_type_refinement
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs
// vybe-test-mode: compile

class Animal {}
class Dog extends Animal {}

abstract class AnimalFactory {
    abstract public function create(): Animal;
}

class DogFactory extends AnimalFactory {
    public function create(): Dog {
        return new Dog();
    }
}

$df = new DogFactory();
echo get_class($df->create());
