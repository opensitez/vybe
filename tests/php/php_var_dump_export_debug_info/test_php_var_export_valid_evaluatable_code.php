<?php
// vybe-test: php/php_var_dump_export_debug_info/test_php_var_export_valid_evaluatable_code
// origin: languages/php/tests/php/test_php_var_dump_export_debug_info.rs

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

$data = ["a" => 1, "b" => [2, 3]];
$code = var_export($data, return: true);
eval('$restored = ' . $code . ';');
echo "a={$restored['a']} b0={$restored['b'][0]}";

__vybe_check(ob_get_clean(), "a=1 b0=2");
