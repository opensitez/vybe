<?php
// vybe-test: php/mbstring/mbencoding_aliases_lists_utf8_aliases
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

$aliases = mb_encoding_aliases('UTF-8');
echo in_array('utf-8', array_map('strtolower', $aliases), true) ? 'alias' : 'none';

__vybe_check(ob_get_clean(), "alias");
