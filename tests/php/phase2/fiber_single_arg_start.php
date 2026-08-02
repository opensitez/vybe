<?php
// vybe-test: php/phase2/fiber_single_arg_start
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

$fiber = new Fiber(function($x) {
    return $x * 2;
});
$result = $fiber->start(21);
