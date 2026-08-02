<?php
// vybe-test: php/pcre_named_groups/positive_lookahead_matches_without_consuming
// origin: languages/php/tests/php/test_pcre_named_groups.rs

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

preg_match_all('/\d+(?= dollars)/', 'I have 50 dollars and 30 euros', $m);
echo implode(',', $m[0]);

__vybe_check(ob_get_clean(), "50");
