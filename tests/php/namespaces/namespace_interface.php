<?php
// vybe-test: php/namespaces/namespace_interface
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Contracts;
interface Renderable {
    public function render(): string;
}

namespace UI;
use Contracts\Renderable;
class Button implements Renderable {
    public function __construct(private string $label) {}
    public function render(): string { return "<button>{$this->label}</button>"; }
}
$btn = new Button("Click me");
echo $btn->render();
