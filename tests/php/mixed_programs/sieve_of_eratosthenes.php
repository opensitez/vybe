<?php
// vybe-test: php/mixed_programs/sieve_of_eratosthenes
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

function sieve(int $limit): array {
    $composite = array_fill(2, $limit - 1, false);
    for ($i = 2; $i * $i <= $limit; $i++) {
        if (!$composite[$i]) {
            for ($j = $i * $i; $j <= $limit; $j += $i) $composite[$j] = true;
        }
    }
    return array_keys(array_filter($composite, fn($v) => !$v));
}
echo implode(',', sieve(30));

__vybe_check(ob_get_clean(), "2,3,5,7,11,13,17,19,23,29");
