<?php
// vybe-test: php/type_juggling/coercion_without_strict_types
// origin: languages/php/tests/php/test_type_juggling.rs

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

// Without strict_types, PHP coerces args
function addNums(int $a, int $b): int { return $a + $b; }
echo addNums("3", "4");  // coerces strings to ints
echo addNums(2.9, 1.1);  // coerces floats to ints (truncates)


__vybe_check(ob_get_clean(), "73");
