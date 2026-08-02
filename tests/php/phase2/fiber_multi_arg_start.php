<?php
// vybe-test: php/phase2/fiber_multi_arg_start
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

$fiber = new Fiber(function($a, $b, $c) {
    return $a + $b + $c;
});
$result = $fiber->start(10, 20, 30);
