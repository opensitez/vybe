<?php
// vybe-test: php/variable_variables/dynamic_method_dispatch
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Math {
    public function double(int $n): int { return $n * 2; }
    public function triple(int $n): int { return $n * 3; }
    public function square(int $n): int { return $n * $n; }
}
$m = new Math();
foreach (['double', 'triple', 'square'] as $op) {
    echo $m->$op(4) . ' ';
}
