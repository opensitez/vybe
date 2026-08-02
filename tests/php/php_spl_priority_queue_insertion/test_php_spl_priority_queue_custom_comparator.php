<?php
// vybe-test: php/php_spl_priority_queue_insertion/test_php_spl_priority_queue_custom_comparator
// origin: languages/php/tests/php/test_php_spl_priority_queue_insertion.rs
// vybe-test-mode: compile

class ReversePriorityQueue extends SplPriorityQueue {
    public function compare(mixed $priority1, mixed $priority2): int {
        return $priority2 <=> $priority1; // Min priority first
    }
}

$pq = new ReversePriorityQueue();
$pq->insert("Urgent", 1);
$pq->insert("Low", 10);
echo $pq->extract() === "Urgent" ? "REVERSE_PRIORITY_OK" : "FAIL";
