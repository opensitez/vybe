<?php
// vybe-test: php/spl_runtime/spl_queue_enqueue_dequeue
// origin: languages/php/tests/php/test_spl_runtime.rs

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

$q = new SplQueue();
$q->enqueue(1);
$q->enqueue(2);
echo $q->dequeue();

__vybe_check(ob_get_clean(), "1");
