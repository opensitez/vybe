<?php
// vybe-test: php/closures_advanced/closure_from_callable_method
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Math {
    public function square(int $n): int { return $n * $n; }
}
$m = new Math();
$fn = Closure::fromCallable([$m, 'square']);
echo $fn(9);
