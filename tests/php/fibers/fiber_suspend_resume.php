<?php
// vybe-test: php/fibers/fiber_suspend_resume
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

$fiber = new Fiber(function() {
    $value = Fiber::suspend('first');
    echo "Resumed with: " . $value;
    Fiber::suspend('second');
});
$v1 = $fiber->start();
echo $v1;
$v2 = $fiber->resume('hello');
echo $v2;
