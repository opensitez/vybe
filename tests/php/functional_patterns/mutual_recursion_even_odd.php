<?php
// vybe-test: php/functional_patterns/mutual_recursion_even_odd
// origin: languages/php/tests/php/test_functional_patterns.rs

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

function isEven(int $n): bool { return $n === 0 ? true : isOdd($n - 1); }
function isOdd(int $n): bool { return $n === 0 ? false : isEven($n - 1); }
echo isEven(4) ? 'even' : 'odd';
echo isOdd(7) ? 'odd' : 'even';

__vybe_check(ob_get_clean(), "evenodd");
