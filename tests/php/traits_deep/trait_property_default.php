<?php
// vybe-test: php/traits_deep/trait_property_default
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait HasName {
    private string $name = 'unnamed';
    public function getName(): string { return $this->name; }
    public function setName(string $n): void { $this->name = $n; }
}
class Animal { use HasName; }
$a = new Animal();
echo $a->getName();
$a->setName('Rex');
echo $a->getName();
