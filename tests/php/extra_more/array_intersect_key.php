<?php
// vybe-test: php/extra_more/array_intersect_key
// origin: languages/php/tests/php/test_extra_more.rs

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

$r=array_intersect_key(['a'=>1,'b'=>2,'c'=>3],['a'=>99,'c'=>99]); echo implode(',',array_keys($r));

__vybe_check(ob_get_clean(), "a,c");
