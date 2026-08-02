<?php
// vybe-test: php/programs/gcd_lcm_calculation
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

function gcd(int $a, int $b): int { return $b === 0 ? $a : gcd($b, $a % $b); }
function lcm(int $a, int $b): int { return intdiv($a * $b, gcd($a, $b)); }
echo gcd(12, 8) . "\n";
echo gcd(48, 18) . "\n";
echo lcm(4, 6) . "\n";
echo lcm(7, 5) . "\n";

__vybe_check(ob_get_clean(), "4\n6\n12\n35");
