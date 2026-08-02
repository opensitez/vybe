<?php
// vybe-test: php/spl_extra/spl_queue_count_changes
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$q = new SplQueue();
echo $q->count();
$q->enqueue('a'); $q->enqueue('b'); $q->enqueue('c');
echo $q->count();
$q->dequeue();
echo $q->count();
