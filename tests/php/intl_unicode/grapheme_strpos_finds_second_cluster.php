<?php
// vybe-test: php/intl_unicode/grapheme_strpos_finds_second_cluster
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

if (!function_exists('grapheme_strpos')) { echo 'skip'; } else {
    echo grapheme_strpos('日本語', '語');
}

__vybe_check(ob_get_clean(), "2");
