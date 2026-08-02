<?php
// vybe-test: php/namespaces/namespace_basic
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace App;
class Greeter {
    public function greet(string $name): string {
        return "Hello, $name!";
    }
}
$g = new Greeter();
echo $g->greet("World");
