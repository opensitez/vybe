<?php
// vybe-test: php/php_spl_queue_fifo_enqueue_dequeue/test_php_spl_queue_fifo_iteration_order
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
$q->enqueue(10);
$q->enqueue(20);
$q->enqueue(30);

$out = [];
foreach ($q as $v) {
    $out[] = $v;
}
echo implode(",", $out);

__vybe_check(ob_get_clean(), "10,20,30");
