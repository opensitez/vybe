<?php
// vybe-test: php/fibers/generator_send_after_start
// origin: languages/php/tests/php/test_fibers.rs

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

function chain() {
    $v = yield 'start';
    yield $v . '-next';
}
$g = chain();
echo $g->current();
echo '|';
echo $g->send('sent');

__vybe_check(ob_get_clean(), "start|sent-next");
