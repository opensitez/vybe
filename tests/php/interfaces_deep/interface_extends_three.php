<?php
// vybe-test: php/interfaces_deep/interface_extends_three
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Named    { public function getName(): string; }
interface Aged     { public function getAge(): int; }
interface Skilled  { public function getSkills(): array; }
interface Person extends Named, Aged, Skilled {}
class Developer implements Person {
    public function __construct(private string $name, private int $age, private array $skills) {}
    public function getName(): string  { return $this->name; }
    public function getAge(): int      { return $this->age; }
    public function getSkills(): array { return $this->skills; }
}
$d = new Developer('Alice', 30, ['PHP', 'Rust']);
echo $d->getName() . ':' . $d->getAge() . ':' . implode(',', $d->getSkills());
