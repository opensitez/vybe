<?php
// vybe-test: php/match_advanced/match_all_arms_with_logical_operators_and_short_circuit
// origin: languages/php/tests/php/test_match_advanced.rs

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

$calls = 0;
$mark = function() use (&$calls) { $calls++; return true; };
$label = match (true) {
    false && $mark() => 'never',
    true && $mark()  => 'hit',
    default => 'none',
};
echo $label;
echo '|';
echo $calls;

__vybe_check(ob_get_clean(), "hit|1");
