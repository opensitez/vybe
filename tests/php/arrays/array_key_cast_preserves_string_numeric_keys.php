<?php
// vybe-test: php/arrays/array_key_cast_preserves_string_numeric_keys
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

$a = ['01' => 'a', 1 => 'b', '2x' => 'c'];
echo array_key_exists('1', $a) ? 'one' : 'no';
echo '|';
echo array_key_exists(1, $a) ? 'num1' : 'no';

__vybe_check(ob_get_clean(), "one|no");
