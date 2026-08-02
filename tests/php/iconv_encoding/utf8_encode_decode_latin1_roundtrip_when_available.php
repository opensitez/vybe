<?php
// vybe-test: php/iconv_encoding/utf8_encode_decode_latin1_roundtrip_when_available
// origin: languages/php/tests/php/test_iconv_encoding.rs

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

if (!function_exists('utf8_encode')) { echo 'skip'; } else {
    echo utf8_decode(utf8_encode("\xE9"));
}

__vybe_check(ob_get_clean(), "é");
