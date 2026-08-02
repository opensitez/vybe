<?php
// vybe-test: php/mbstring/mbstrpad_pads_to_width_when_available
// origin: languages/php/tests/php/test_mbstring.rs

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

if (function_exists('mb_str_pad')) {
    echo mb_str_pad('hi', 5, '0', STR_PAD_LEFT);
} else {
    echo '00hi';
}

__vybe_check(ob_get_clean(), "000hi");
