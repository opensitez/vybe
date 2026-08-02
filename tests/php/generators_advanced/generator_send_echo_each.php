<?php
// vybe-test: php/generators_advanced/generator_send_echo_each
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function logger() {
    while (true) {
        $msg = yield;
        if ($msg === null) break;
        echo strtoupper($msg);
    }
}
$log = logger();
$log->current();
$log->send("hello");
$log->send("world");

__vybe_check(ob_get_clean(), "HELLOWORLD");
