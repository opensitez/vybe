<?php
// vybe-test: php/match_edge_cases/match_default_arm_catches_all_unmatched
// origin: languages/php/tests/php/test_match_edge_cases.rs

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

for ($i = 1; $i <= 3; $i++) {
    echo match($i) { 1 => 'one', default => 'many' } . ',';
}

__vybe_check(ob_get_clean(), "one,many,many,");
