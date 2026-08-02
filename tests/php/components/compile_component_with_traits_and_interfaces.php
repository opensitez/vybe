<?php
// vybe-test: php/components/compile_component_with_traits_and_interfaces
// origin: languages/php/tests/php/test_components.rs
// vybe-test-mode: compile

interface Printable {
    public function toString(): string;
}
trait Loggable {
    public function log() { echo $this->toString(); }
}
class Item implements Printable {
    use Loggable;
    public $name;
    public function __construct($name) { $this->name = $name; }
    public function toString(): string { return $this->name; }
}
