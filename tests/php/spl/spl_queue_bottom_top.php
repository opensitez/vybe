<?php
// vybe-test: php/spl/spl_queue_bottom_top
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$q = new SplQueue();
$q->enqueue('first');
$q->enqueue('middle');
$q->enqueue('last');
echo $q->bottom() . ',' . $q->top();
