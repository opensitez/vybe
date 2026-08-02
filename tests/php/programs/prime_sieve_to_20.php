<?php
// vybe-test: php/programs/prime_sieve_to_20
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

function sieve(int $n): array {
    $is_prime = array_fill(2, $n - 1, true);
    for ($i = 2; $i * $i <= $n; $i++) {
        if ($is_prime[$i]) {
            for ($j = $i * $i; $j <= $n; $j += $i) $is_prime[$j] = false;
        }
    }
    return array_keys(array_filter($is_prime));
}
echo implode(',', sieve(20)) . "\n";

__vybe_check(ob_get_clean(), "2,3,5,7,11,13,17,19");
