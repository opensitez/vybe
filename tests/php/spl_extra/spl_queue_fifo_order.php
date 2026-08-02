<?php
// vybe-test: php/spl_extra/spl_queue_fifo_order
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$q = new SplQueue();
foreach (['x', 'y', 'z'] as $v) { $q->enqueue($v); }
$out = [];
while (!$q->isEmpty()) { $out[] = $q->dequeue(); }
echo implode(',', $out); // x,y,z
