<?php
// vybe-test: php/php_array_uassoc_callback_key_value/test_array_uassoc_custom_key_comparison
// origin: languages/php/tests/php/test_php_array_uassoc_callback_key_value.rs

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

$a1 = ['a' => 1, 'b' => 2];
$a2 = ['A' => 1, 'B' => 3];
$diff = array_uassoc($a1, $a2, 'strcasecmp');
echo count($diff) . ':' . implode(',', array_keys($diff)), "\n";

__vybe_check(ob_get_clean(), "1:b");
