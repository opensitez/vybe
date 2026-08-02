<?php
// vybe-test: php/php_spl_queue_fifo_enqueue_dequeue/test_php_spl_queue_fifo_enqueue_and_dequeue
// origin: languages/php/tests/php/test_php_spl_queue_fifo_enqueue_dequeue.rs

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
$q->enqueue("first_in");
$q->enqueue("second_in");

echo $q->dequeue() . " | " . $q->dequeue();

__vybe_check(ob_get_clean(), "first_in | second_in");
