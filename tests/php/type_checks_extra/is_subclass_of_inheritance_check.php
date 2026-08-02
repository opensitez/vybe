<?php
// vybe-test: php/type_checks_extra/is_subclass_of_inheritance_check
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

class Animal {}
class Dog extends Animal {}
$d = new Dog();
echo is_subclass_of($d, 'Animal') ? 'yes' : 'no';
echo is_subclass_of($d, 'Dog') ? 'yes' : 'no';
