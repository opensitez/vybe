<?php
// vybe-test: php/arrays/array_destructure_nested_and_default
// origin: languages/php/tests/php/test_arrays.rs

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

$data = [['a', 'b'], ['c']];
[$left, $right] = $data;
[$x, $y] = $left;
// PHP list()/[] destructuring has no default syntax — `[$a, $b = 'x'] = …` is
// "Assignments can only happen to writable values". The default is spelled with
// `??` on the read.
[$x2] = $right;
$y2 = $right[1] ?? 'zz';
echo $x . $y . $x2 . $y2;

__vybe_check(ob_get_clean(), "abczz");
