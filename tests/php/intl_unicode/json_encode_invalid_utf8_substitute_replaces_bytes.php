<?php
// vybe-test: php/intl_unicode/json_encode_invalid_utf8_substitute_replaces_bytes
// origin: languages/php/tests/php/test_intl_unicode.rs

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

$out = json_encode("\xB1\x31", JSON_INVALID_UTF8_SUBSTITUTE);
echo str_contains($out, '1') ? 'substituted' : 'lost';

__vybe_check(ob_get_clean(), "substituted");
