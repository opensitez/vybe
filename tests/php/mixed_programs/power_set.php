<?php
// vybe-test: php/mixed_programs/power_set
// origin: languages/php/tests/php/test_mixed_programs.rs

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

function powerSet(array $set): array {
    if (!$set) return [[]];
    $first = array_shift($set);
    $rest = powerSet($set);
    return array_merge($rest, array_map(fn($s) => array_merge([$first], $s), $rest));
}
$ps = powerSet([1,2,3]);
echo count($ps);

__vybe_check(ob_get_clean(), "8");
