<?php
// vybe-test: php/traits_deep/trait_abstract_requirement
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait Printable {
    abstract public function toString(): string;
    public function print(): void { echo $this->toString(); }
}
class Color {
    use Printable;
    public function __construct(private string $name) {}
    public function toString(): string { return "Color({$this->name})"; }
}
(new Color('red'))->print();
