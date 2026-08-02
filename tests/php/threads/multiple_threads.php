<?php
// vybe-test: php/threads/multiple_threads
// origin: languages/php/tests/php/test_threads.rs
// vybe-test-mode: compile

$t1 = thread_create(fn() => 1 + 2);
$t2 = thread_create(fn() => 3 + 4);
$r1 = thread_join($t1);
$r2 = thread_join($t2);
echo $r1 + $r2;
