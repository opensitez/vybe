<?php
// vybe-test: php/interfaces_deep/covariant_return_self_static
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Buildable { public function withName(string $name): static; }
class Widget implements Buildable {
    private string $name = '';
    public function withName(string $name): static {
        $clone = clone $this;
        $clone->name = $name;
        return $clone;
    }
    public function getName(): string { return $this->name; }
}
class Button extends Widget {}
$btn = (new Button())->withName('Submit');
echo $btn->getName();
echo ($btn instanceof Button) ? ':is Button' : ':not Button';
