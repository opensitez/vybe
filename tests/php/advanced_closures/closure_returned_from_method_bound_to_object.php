<?php
// vybe-test: php/advanced_closures/closure_returned_from_method_bound_to_object
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

class Greeter {
    private string $prefix;
    public function __construct(string $prefix) { $this->prefix = $prefix; }
    public function makeGreeter(): Closure {
        return function(string $name): string { return $this->prefix . $name; };
    }
}
$g = new Greeter('Hello, ');
$fn = $g->makeGreeter();
echo $fn('World');
