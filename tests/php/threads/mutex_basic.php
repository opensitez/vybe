<?php
// vybe-test: php/threads/mutex_basic
// origin: languages/php/tests/php/test_threads.rs
// vybe-test-mode: compile

$lock = mutex_create();
mutex_lock($lock);
$shared = 42;
mutex_unlock($lock);
