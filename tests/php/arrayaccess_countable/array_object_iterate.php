<?php
// vybe-test: php/arrayaccess_countable/array_object_iterate
// origin: languages/php/tests/php/test_arrayaccess_countable.rs

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

$ao = new ArrayObject(['a' => 1, 'b' => 2, 'c' => 3]);
foreach ($ao as $k => $v) echo $k . $v;

__vybe_check(ob_get_clean(), "a1b2c3");
