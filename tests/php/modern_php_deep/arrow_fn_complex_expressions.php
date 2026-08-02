<?php
// vybe-test: php/modern_php_deep/arrow_fn_complex_expressions
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

$add     = fn(int $a, int $b): int => $a + $b;
$compose = fn(callable $f, callable $g) => fn($x) => $f($g($x));
$double  = fn($x) => $x * 2;
$inc     = fn($x) => $x + 1;
$doubleInc = $compose($double, $inc);
echo $add(3, 4);
echo $doubleInc(5);

__vybe_check(ob_get_clean(), "712");
