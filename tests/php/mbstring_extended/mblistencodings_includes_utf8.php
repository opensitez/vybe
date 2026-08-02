<?php
// vybe-test: php/mbstring_extended/mblistencodings_includes_utf8
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

echo in_array('UTF-8', mb_list_encodings(), true) ? 'utf8' : 'missing';

__vybe_check(ob_get_clean(), "utf8");
