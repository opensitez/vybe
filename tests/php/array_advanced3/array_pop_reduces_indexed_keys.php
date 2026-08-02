<?php
// vybe-test: php/array_advanced3/array_pop_reduces_indexed_keys
// origin: languages/php/tests/php/test_array_advanced3.rs

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

$a = ['a' => 1, 1 => 2, 4 => 3];
$last = array_pop($a);
echo $last;
echo '|';
echo json_encode(array_keys($a));
echo '|';
echo isset($a[4]) ? 'has4' : 'no4';

__vybe_check(ob_get_clean(), "3|[\"a\",1]|no4");
