<?php
// vybe-test: php/array_search_sort/array_key_exists_on_numeric_string_key_collisions
// origin: languages/php/tests/php/test_array_search_sort.rs

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

$a = ['0' => 'zero', 0 => 'int0', '01' => 'leading'];
echo array_key_exists(0, $a) . '|' . array_key_exists('0', $a) . '|' . array_key_exists('01', $a);

__vybe_check(ob_get_clean(), "1|1|1");
