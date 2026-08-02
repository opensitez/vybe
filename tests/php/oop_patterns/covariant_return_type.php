<?php
// vybe-test: php/oop_patterns/covariant_return_type
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Animal {}
class Dog extends Animal {}
class AnimalFactory {
    public function create(): Animal { return new Animal(); }
}
class DogFactory extends AnimalFactory {
    public function create(): Dog { return new Dog(); }
}
$f = new DogFactory();
echo $f->create() instanceof Dog ? 'dog' : 'not-dog';
