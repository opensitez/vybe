<?php
// vybe-test: php/enums/enum_in_array_keyed_by_name
// origin: languages/php/tests/php/test_enums.rs

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

enum Tier: string { case Free = 'free'; case Pro = 'pro'; }
$map = [Tier::Free->name => 0, Tier::Pro->name => 1];
echo $map['Pro'];

__vybe_check(ob_get_clean(), "1");
