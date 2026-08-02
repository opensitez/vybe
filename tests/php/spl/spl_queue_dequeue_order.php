<?php
// vybe-test: php/spl/spl_queue_dequeue_order
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$q = new SplQueue();
foreach (['a', 'b', 'c'] as $v) { $q->enqueue($v); }
$result = [];
while (!$q->isEmpty()) { $result[] = $q->dequeue(); }
echo implode(',', $result);
