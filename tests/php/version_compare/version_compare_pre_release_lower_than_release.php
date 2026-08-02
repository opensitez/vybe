<?php
// vybe-test: php/version_compare/version_compare_pre_release_lower_than_release
// origin: languages/php/tests/php/test_version_compare.rs

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

echo version_compare('8.2.0beta1', '8.2.0', '<') ? 'beta' : 'rel';

__vybe_check(ob_get_clean(), "beta");
