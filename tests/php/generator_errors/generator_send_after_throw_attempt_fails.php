<?php
// vybe-test: php/generator_errors/generator_send_after_throw_attempt_fails
// origin: languages/php/tests/php/test_generator_errors.rs

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

function acc(): Generator {
    $total = 0;
    while (true) {
        $n = yield $total;
        if ($n === null) break;
        $total += $n;
    }
}
$g = acc();
$g->current();
try {
    $g->throw(new InvalidArgumentException('abort'));
    $g->send(5);
    echo 'sent';
} catch (InvalidArgumentException $e) {
    echo 'abort';
}

__vybe_check(ob_get_clean(), "abort");
