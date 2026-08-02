<?php
// vybe-test: php/arrays/array_search_and_replace_with_missing_key
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

$a = ['a' => 1, 'b' => 2];
echo array_key_exists('z', $a) ? 'yes' : 'no';
echo '|';
echo array_search(3, $a, true) === false ? 'missing' : 'found';

__vybe_check(ob_get_clean(), "no|missing");
