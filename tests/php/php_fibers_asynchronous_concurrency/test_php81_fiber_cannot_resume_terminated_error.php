<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_cannot_resume_terminated_error
// origin: languages/php/tests/php/test_php_fibers_asynchronous_concurrency.rs
// vybe-test-mode: compile

$fiber = new Fiber(fn() => 123);
$fiber->start();
try {
    $fiber->resume();
} catch (FiberError $e) {
    echo "Cannot resume terminated fiber";
}
