<?php
// vybe-test: php/new_features/interface_decl
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

interface Printable {
    public function toString();
}
class Item implements Printable {
    public $name;
    public function __construct($n) { $this->name = $n; }
    public function toString() { return $this->name; }
}
$i = new Item("test");
echo $i->toString();
