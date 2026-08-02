<?php
// vybe-test: php/threads/fiber_in_thread
// origin: languages/php/tests/php/test_threads.rs
// vybe-test-mode: compile

$thread = thread_create(function() {
    $fiber = new Fiber(function() {
        Fiber::suspend('from thread fiber');
        return 'done';
    });
    $v = $fiber->start();
    $fiber->resume();
    return $fiber->getReturn();
});
$result = thread_join($thread);
