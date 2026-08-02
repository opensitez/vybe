<?php
// vybe-test: php/spl_object_storage/spl_priority_queue_order
// origin: languages/php/tests/php/test_spl_object_storage.rs

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

$pq = new SplPriorityQueue;
$pq->insert('low', 1);
$pq->insert('high', 3);
$pq->insert('mid', 2);
$out = [];
while (!$pq->isEmpty()) $out[] = $pq->extract();
echo implode(',', $out);

__vybe_check(ob_get_clean(), "high,mid,low");
