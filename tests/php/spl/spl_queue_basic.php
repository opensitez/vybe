<?php
// vybe-test: php/spl/spl_queue_basic
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$q = new SplQueue();
$q->enqueue('first');
$q->enqueue('second');
$q->enqueue('third');
echo $q->count();
