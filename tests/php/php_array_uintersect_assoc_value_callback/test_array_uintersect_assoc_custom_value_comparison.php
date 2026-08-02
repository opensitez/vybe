<?php
// vybe-test: php/php_array_uintersect_assoc_value_callback/test_array_uintersect_assoc_custom_value_comparison
// origin: languages/php/tests/php/test_php_array_uintersect_assoc_value_callback.rs

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

$a1 = ['x' => '10', 'y' => '20'];
$a2 = ['x' => 10, 'y' => 30];
$intersection = array_uintersect_assoc($a1, $a2, function($v1, $v2) {
    return (int)$v1 <=> (int)$v2;
});
echo count($intersection) . ':' . implode(',', array_keys($intersection)), "\n";

__vybe_check(ob_get_clean(), "1:x");
