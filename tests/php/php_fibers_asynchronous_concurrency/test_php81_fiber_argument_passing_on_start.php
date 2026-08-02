<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_argument_passing_on_start
// origin: languages/php/tests/php/test_php_fibers_asynchronous_concurrency.rs
// vybe-test-mode: compile

$fiber = new Fiber(function(string $name, int $id): string {
    return "$name#$id";
});
$res = $fiber->start("Worker", 99);
echo $res;
