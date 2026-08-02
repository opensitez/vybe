<?php
// vybe-test: php/fibers/fiber_get_return
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

$fiber = new Fiber(function() {
    return 'done';
});
$fiber->start();
$result = $fiber->getReturn();
