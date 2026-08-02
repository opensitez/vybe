<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_current_reference
// origin: languages/php/tests/php/test_php_fibers_asynchronous_concurrency.rs
// vybe-test-mode: compile

$fiber = new Fiber(function(): void {
    $curr = Fiber::getCurrent();
    echo $curr !== null ? "HAS_CURRENT" : "NO_CURRENT";
});
$fiber->start();
