<?php
// vybe-test: php/operators/compound_assignment_operator_runtime_results
// origin: languages/php/tests/php/test_operators.rs

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

$x = 5;
$x += 2;
echo $x;
$x -= 4;
echo $x;
$x *= 3;
echo $x;
$x /= 9;
echo $x;
$x %= 2;
echo $x;

$text = 'a';
$text .= 'b';
echo $text;

$bits = 6;
$bits &= 3;
echo $bits;
$bits |= 4;
echo $bits;
$bits ^= 1;
echo $bits;

$shift = 1;
$shift <<= 3;
echo $shift;
$shift >>= 2;
echo $shift;

$fallback = null;
$fallback ??= 'set';
echo $fallback;
$fallback ??= 'again';
echo $fallback;

__vybe_check(ob_get_clean(), "73911ab26782setset");
