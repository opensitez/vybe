<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_spl_queue_enqueue_dequeue_fifo
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs

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

$queue = new SplQueue();
$queue->enqueue("A");
$queue->enqueue("B");
echo $queue->dequeue() . " -> " . $queue->dequeue();

__vybe_check(ob_get_clean(), "A -> B");
