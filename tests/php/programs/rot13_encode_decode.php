<?php
// vybe-test: php/programs/rot13_encode_decode
// origin: languages/php/tests/php/test_programs.rs

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

function rot13(string $s): string { return str_rot13($s); }
$encoded = rot13('Hello World');
echo $encoded . "\n";
echo rot13($encoded) . "\n";

__vybe_check(ob_get_clean(), "Uryyb Jbeyq\nHello World");
