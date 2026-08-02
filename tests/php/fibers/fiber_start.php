<?php
// vybe-test: php/fibers/fiber_start
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

$fiber = new Fiber(function() {
    echo "Running";
    return 42;
});
$result = $fiber->start();
