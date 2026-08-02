<?php
// vybe-test: php/mb_stristr_case_folding/mb_stristr_basic
// origin: languages/php/tests/php/test_mb_stristr_case_folding.rs

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

$str = "ÄÖÜ test";
echo mb_stristr($str, "öü", true, "UTF-8");

__vybe_check(ob_get_clean(), "Ä");
