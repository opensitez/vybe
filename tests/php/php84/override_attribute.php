<?php
// vybe-test: php/php84/override_attribute
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class Base { public function render(): string { return 'base'; } }
class Derived extends Base {
    #[\Override]
    public function render(): string { return 'derived'; }
}
echo (new Derived())->render();
