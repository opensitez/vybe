<?php
// vybe-test: php/interfaces_deep/covariant_return_type_basic
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

class Animal {}
class Dog extends Animal {}
interface AnimalFactory { public function create(): Animal; }
class DogFactory implements AnimalFactory {
    public function create(): Dog { return new Dog(); }
}
$factory = new DogFactory();
$dog = $factory->create();
echo ($dog instanceof Animal) ? 'is Animal' : 'not Animal';
echo ($dog instanceof Dog) ? ':is Dog' : ':not Dog';
