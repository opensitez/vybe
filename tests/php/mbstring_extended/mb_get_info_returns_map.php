<?php
// vybe-test: php/mbstring_extended/mb_get_info_returns_map
// origin: languages/php/tests/php/test_mbstring_extended.rs

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

$info = mb_get_info();
echo is_array($info) ? 'array' : 'none';
echo '|';
echo $info['internal_encoding'] === mb_internal_encoding() ? 'ok' : 'bad';

__vybe_check(ob_get_clean(), "array|ok");
