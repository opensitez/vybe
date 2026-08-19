<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_array_search_and_flip_reduction
// origin: languages/php/tests/php/test_php_array_key_operations_indexing.rs

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

$roles = ['read', 'write', 'read', 'admin'];
$pos = array_search('admin', $roles, true);
$flipped = array_flip($roles);
echo $pos . '|';
echo $flipped['read'];

__vybe_check(ob_get_clean(), "3|2");
