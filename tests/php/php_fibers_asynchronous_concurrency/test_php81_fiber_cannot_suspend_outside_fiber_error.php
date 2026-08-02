<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_cannot_suspend_outside_fiber_error
// origin: languages/php/tests/php/test_php_fibers_asynchronous_concurrency.rs
// vybe-test-mode: compile

try {
    Fiber::suspend();
} catch (FiberError $e) {
    echo "FiberError: " . $e->getMessage();
}
