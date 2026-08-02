<?php
// vybe-test: php/closures_advanced/closure_bind_to_basic
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Box { private int $size = 10; }
$fn = function() { return $this->size; };
$bound = $fn->bindTo(new Box(), Box::class);
echo $bound();
