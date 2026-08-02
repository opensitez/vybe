<?php
// vybe-test: php/closures_advanced/closure_bind_basic
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Foo { private int $x = 42; }
$fn = Closure::bind(function() { return $this->x; }, new Foo(), 'Foo');
echo $fn();
