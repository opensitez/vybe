<?php
// vybe-test: php/arrays/array_search_with_numeric_like_string_keys_runtime
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

$a = ['01' => 'left', '1' => 'right', 2 => 'third'];
echo $a['1'];
echo '|';
echo $a['01'];
echo '|';
echo array_key_exists(1, $a) ? 'one' : 'no';

__vybe_check(ob_get_clean(), "right|left|one");
