<?php
// vybe-test: php/phase2/fiber_no_arg_start
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

$fiber = new Fiber(function() {
    return 42;
});
$result = $fiber->start();
