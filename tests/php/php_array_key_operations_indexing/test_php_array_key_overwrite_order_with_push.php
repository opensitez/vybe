<?php
// vybe-test: php/php_array_key_operations_indexing/test_php_array_key_overwrite_order_with_push
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

$a = [];
$a['x'] = 1;
$a[0] = 'zero';
$a[] = 'append';
echo count($a) . ':' . implode('|', $a);

__vybe_check(ob_get_clean(), "3:1|zero|append");
