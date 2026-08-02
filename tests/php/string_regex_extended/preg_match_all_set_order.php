<?php
// vybe-test: php/string_regex_extended/preg_match_all_set_order
// origin: languages/php/tests/php/test_string_regex_extended.rs

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

preg_match_all('/(\w+):(\d+)/', 'foo:1 bar:2', $m, PREG_SET_ORDER);
echo $m[0][1] . '=' . $m[0][2] . ',' . $m[1][1] . '=' . $m[1][2];
echo "\n";

__vybe_check(ob_get_clean(), "foo=1,bar=2");
