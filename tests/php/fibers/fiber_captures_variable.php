<?php
// vybe-test: php/fibers/fiber_captures_variable
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

$message = "Hello";
$fiber = new Fiber(function() use ($message) {
    Fiber::suspend($message . " World");
});
$result = $fiber->start();
echo $result;
