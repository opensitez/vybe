<?php
// vybe-test: php/string_escaping/quoted_printable_encode_decode_roundtrip
// origin: languages/php/tests/php/test_string_escaping.rs

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

$qp = quoted_printable_encode("a=b");
echo quoted_printable_decode($qp);

__vybe_check(ob_get_clean(), "a=b");
