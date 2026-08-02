<?php
// vybe-test: php/fibers/fiber_with_args
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

$fiber = new Fiber(function($x, $y) {
    Fiber::suspend($x + $y);
    return $x * $y;
});
$sum = $fiber->start(3, 4);
echo $sum;
