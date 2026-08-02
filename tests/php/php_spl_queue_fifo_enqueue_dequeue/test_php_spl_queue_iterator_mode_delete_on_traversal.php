<?php
// vybe-test: php/php_spl_queue_fifo_enqueue_dequeue/test_php_spl_queue_iterator_mode_delete_on_traversal
// origin: languages/php/tests/php/test_php_spl_queue_fifo_enqueue_dequeue.rs
// vybe-test-mode: compile

$q = new SplQueue();
$q->enqueue("msg1");
$q->enqueue("msg2");
$q->setIteratorMode(SplQueue::IT_MODE_DELETE);
foreach ($q as $msg) {}
echo count($q) === 0 ? "TRAVERSAL_DELETE_OK" : "FAIL";
