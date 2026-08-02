<?php
// vybe-test: php/oop/interface_impl
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

interface Printable { public function toString(): string; }
class Item implements Printable {
    public $name;
    public function __construct($n) { $this->name = $n; }
    public function toString(): string { return $this->name; }
}
$i = new Item('Widget');
echo $i->toString();
