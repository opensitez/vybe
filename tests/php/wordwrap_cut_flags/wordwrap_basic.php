<?php
// vybe-test: php/wordwrap_cut_flags/wordwrap_basic
// origin: languages/php/tests/php/test_wordwrap_cut_flags.rs

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

$str = "The quick brown fox jumped over the lazy dog.";
echo wordwrap($str, 20, "<br>\n");

__vybe_check(ob_get_clean(), "The quick brown fox<br>\njumped over the lazy<br>\ndog.");
