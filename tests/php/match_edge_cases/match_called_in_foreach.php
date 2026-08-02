<?php
// vybe-test: php/match_edge_cases/match_called_in_foreach
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

$grades = ['A', 'B', 'C', 'F'];
$labels = array_map(fn($g) => match($g) {
    'A' => 'excellent',
    'B' => 'good',
    'C' => 'average',
    default => 'fail',
}, $grades);
echo implode(',', $labels);

__vybe_check(ob_get_clean(), "excellent,good,average,fail");
