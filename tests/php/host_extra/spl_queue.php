<?php
// vybe-test: php/host_extra/spl_queue
// origin: languages/php/tests/php/test_host_extra.rs
// vybe-test-mode: compile

$queue = new SplQueue();
$queue->enqueue('first');
$queue->enqueue('second');
$item = $queue->dequeue();
echo $item;
$next = $queue->peek();
