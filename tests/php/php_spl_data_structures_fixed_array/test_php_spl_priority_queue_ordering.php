<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_spl_priority_queue_ordering
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs
// vybe-test-mode: compile

$pq = new SplPriorityQueue();
$pq->insert("low priority task", 1);
$pq->insert("high priority task", 100);
$pq->insert("medium priority task", 50);

while ($pq->valid()) {
    echo $pq->current() . "\n";
    $pq->next();
}
