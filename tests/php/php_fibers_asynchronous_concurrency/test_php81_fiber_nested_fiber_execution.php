<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_nested_fiber_execution
// origin: languages/php/tests/php/test_php_fibers_asynchronous_concurrency.rs
// vybe-test-mode: compile

$f1 = new Fiber(function(): void {
    $f2 = new Fiber(function(): void {
        Fiber::suspend("inner_yield");
    });
    $val = $f2->start();
    echo "F1 received $val";
});
$f1->start();
