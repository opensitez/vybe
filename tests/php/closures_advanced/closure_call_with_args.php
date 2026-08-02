<?php
// vybe-test: php/closures_advanced/closure_call_with_args
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Multiplier { private int $factor = 3; }
$fn = function(int $n): int { return $this->factor * $n; };
echo $fn->call(new Multiplier(), 7);
