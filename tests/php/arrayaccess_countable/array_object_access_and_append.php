<?php
// vybe-test: php/arrayaccess_countable/array_object_access_and_append
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

$ao = new ArrayObject(['x' => 1]);
$ao['y'] = 2;
$ao->append(3);
echo $ao['x'] . ',' . $ao['y'] . ',' . $ao->count();

__vybe_check(ob_get_clean(), "1,2,3");
