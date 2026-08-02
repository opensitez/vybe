<?php
// vybe-test: php/usort/usort_stable_like_behavior_with_tiebreaker
// origin: languages/php/tests/php/test_usort.rs

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

$rows = [['k' => 2, 'n' => 'b'], ['k' => 1, 'n' => 'a'], ['k' => 2, 'n' => 'c']];
usort($rows, function ($a, $b) {
    return $a['k'] <=> $b['k'] ?: $a['n'] <=> $b['n'];
});
echo $rows[0]['n'] . $rows[2]['n'];

__vybe_check(ob_get_clean(), "ac");
