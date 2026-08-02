<?php
// vybe-test: php/spl_extra/spl_queue_enqueue_dequeue
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$q = new SplQueue();
$q->enqueue('alpha');
$q->enqueue('beta');
$q->enqueue('gamma');
echo $q->dequeue();
echo $q->count();
