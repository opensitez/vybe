<?php
// vybe-test: php/variable_variables/dynamic_class_new
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Dog  { public function speak(): string { return "Woof"; } }
class Cat  { public function speak(): string { return "Meow"; } }
class Bird { public function speak(): string { return "Tweet"; } }
foreach (['Dog', 'Cat', 'Bird'] as $cls) {
    $obj = new $cls();
    echo $obj->speak() . ' ';
}
