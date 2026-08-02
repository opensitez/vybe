<?php
// vybe-test: php/enums_deep/backed_enum_with_trait
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

trait Describable {
    public function describe(): string {
        return "{$this->name}={$this->value}";
    }
}
enum Color: string {
    use Describable;
    case Red   = 'red';
    case Green = 'green';
    case Blue  = 'blue';
}
echo Color::Red->describe();
