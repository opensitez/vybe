<?php
// vybe-test: php/fibers/multiple_fibers
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

$f1 = new Fiber(function() {
    Fiber::suspend('f1');
    return 'f1 done';
});
$f2 = new Fiber(function() {
    Fiber::suspend('f2');
    return 'f2 done';
});
echo $f1->start();
echo $f2->start();
$f1->resume();
$f2->resume();
