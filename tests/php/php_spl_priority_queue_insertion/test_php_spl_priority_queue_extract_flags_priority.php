<?php
// vybe-test: php/php_spl_priority_queue_insertion/test_php_spl_priority_queue_extract_flags_priority
// origin: languages/php/tests/php/test_php_spl_priority_queue_insertion.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$pq = new SplPriorityQueue();
$pq->setExtractFlags(SplPriorityQueue::EXTR_PRIORITY);
$pq->insert("Task A", 25);

echo "Priority=" . $pq->extract();

__vybe_check(ob_get_clean(), "Priority=25");
