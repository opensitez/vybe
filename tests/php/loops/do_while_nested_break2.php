<?php
// vybe-test: php/loops/do_while_nested_break2
// origin: languages/php/tests/php/test_loops.rs

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

$out = '';
$outer = 0;
do {
    $inner = 0;
    $outer++;
    do {
        $inner++;
        if ($outer === 2 && $inner === 2) {
            break 2;
        }
        $out .= $inner;
    } while ($inner < 3);
} while ($outer < 4);
echo $out;

__vybe_check(ob_get_clean(), "1231");
