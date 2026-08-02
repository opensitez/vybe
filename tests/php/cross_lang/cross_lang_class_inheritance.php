<?php
// vybe-test: php/cross_lang/cross_lang_class_inheritance
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

// PHP classes use same object layout as Python/JS/VB/C# classes:
// - emit_new_typed_object (same __type stamp)
// - emit_bind_method_with_aliases (cross-lang method names)
// - emit_store_super (__super chain)
// - register_type (type table entry)
// This means a PHP class can extend a Python class at runtime.
class Animal {
    public $name;
    public function __construct($name) { $this->name = $name; }
    public function toString() { return $this->name; }
}
class Dog extends Animal {
    public function speak() { return $this->name . ' barks'; }
}
$d = new Dog('Rex');
echo $d->speak();
echo $d->toString(); // Also callable as __str__() from Python
