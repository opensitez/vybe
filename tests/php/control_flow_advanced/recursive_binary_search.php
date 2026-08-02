<?php
// vybe-test: php/control_flow_advanced/recursive_binary_search
// origin: languages/php/tests/php/test_control_flow_advanced.rs

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

function bsearch(array $a, int $target, int $lo = 0, ?int $hi = null): int {
    $hi ??= count($a) - 1;
    if ($lo > $hi) return -1;
    $mid = intdiv($lo + $hi, 2);
    return match(true) {
        $a[$mid] === $target => $mid,
        $a[$mid] < $target  => bsearch($a, $target, $mid + 1, $hi),
        default              => bsearch($a, $target, $lo, $mid - 1),
    };
}
$sorted = range(0, 20, 2);
echo bsearch($sorted, 14);

__vybe_check(ob_get_clean(), "7");
