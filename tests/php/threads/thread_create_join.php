<?php
// vybe-test: php/threads/thread_create_join
// origin: languages/php/tests/php/test_threads.rs
// vybe-test-mode: compile

$thread = thread_create(function() {
    return 42;
});
$result = thread_join($thread);
echo $result;
