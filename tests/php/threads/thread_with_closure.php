<?php
// vybe-test: php/threads/thread_with_closure
// origin: languages/php/tests/php/test_threads.rs
// vybe-test-mode: compile

$data = "hello";
$thread = thread_create(function() use ($data) {
    return strtoupper($data);
});
$result = thread_join($thread);
