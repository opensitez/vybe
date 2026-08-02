<?php
// vybe-test: php/spl/spl_priority_queue_extract_priority_only_runtime
// origin: languages/php/tests/php/test_spl.rs

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
$pq->insert('alpha', 5);
$pq->insert('beta', 10);
echo $pq->extract();
echo '|';
echo $pq->count();

__vybe_check(ob_get_clean(), "10|1");
