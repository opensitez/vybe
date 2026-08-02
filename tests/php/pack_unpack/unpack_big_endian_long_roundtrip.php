<?php
// vybe-test: php/pack_unpack/unpack_big_endian_long_roundtrip
// origin: languages/php/tests/php/test_pack_unpack.rs

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

$p = pack('N', 305419896);
echo unpack('N', $p)[1];

__vybe_check(ob_get_clean(), "305419896");
