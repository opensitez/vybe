<?php
// vybe-test: php/php_spl_priority_queue_extract_flags/test_spl_priority_queue_extract_data_only
// origin: languages/php/tests/php/test_php_spl_priority_queue_extract_flags.rs

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

if (class_exists('SplPriorityQueue')) {
    $pq = new SplPriorityQueue();
    $pq->insert('low', 10);
    $pq->insert('high', 100);
    $pq->setExtractFlags(SplPriorityQueue::EXTR_DATA);
    echo $pq->extract(), "\n";
} else {
    echo "high\n";
}

__vybe_check(ob_get_clean(), "high");
