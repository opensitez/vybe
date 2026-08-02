<?php
// vybe-test: php/php_spl_queue_fifo_enqueue_dequeue/test_php_spl_queue_push_and_pop_aliases
// origin: languages/php/tests/php/test_php_spl_queue_fifo_enqueue_dequeue.rs
// vybe-test-mode: compile

$q = new SplQueue();
$q->push("a");
$q->push("b");
echo $q->pop() === "b" ? "POP_OK" : "FAIL";
