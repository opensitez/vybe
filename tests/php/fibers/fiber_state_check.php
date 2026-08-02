<?php
// vybe-test: php/fibers/fiber_state_check
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

$fiber = new Fiber(function() {
    Fiber::suspend();
});
$fiber->start();
$suspended = $fiber->isSuspended();
