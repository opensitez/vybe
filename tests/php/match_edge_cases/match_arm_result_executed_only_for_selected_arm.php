<?php
// vybe-test: php/match_edge_cases/match_arm_result_executed_only_for_selected_arm
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

$hits = [];
$result = match (3) {
    1 => (fn() use (&$hits) { $hits[] = 'a'; return 'one'; })(),
    2 => (function() use (&$hits) { $hits[] = 'b'; return 'two'; })(),
    3 => 'three',
    default => (function() use (&$hits) { $hits[] = 'd'; return 'other'; })(),
};
echo $result;
echo '|';
echo implode('|', $hits);

__vybe_check(ob_get_clean(), "three|");
