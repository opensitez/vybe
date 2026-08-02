<?php
// vybe-test: php/programs/memoized_fibonacci
// origin: languages/php/tests/php/test_programs.rs

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

function memoFib(int $n, array &$memo = []): int {
    if ($n <= 1) return $n;
    if (isset($memo[$n])) return $memo[$n];
    $memo[$n] = memoFib($n - 1, $memo) + memoFib($n - 2, $memo);
    return $memo[$n];
}
$m = [];
echo memoFib(10, $m) . "\n";
echo memoFib(20, $m) . "\n";

__vybe_check(ob_get_clean(), "55\n6765");
