<?php
// vybe-test: php/threads/mutex_in_threads
// origin: languages/php/tests/php/test_threads.rs
// vybe-test-mode: compile

$lock = mutex_create();
$counter = 0;

$t1 = thread_create(function() use ($lock) {
    mutex_lock($lock);
    mutex_unlock($lock);
});

$t2 = thread_create(function() use ($lock) {
    mutex_lock($lock);
    mutex_unlock($lock);
});

thread_join($t1);
thread_join($t2);
