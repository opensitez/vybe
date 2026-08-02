<?php
// vybe-test: php/php_extension_funcs_loaded_list/test_get_extension_funcs_nonexistent
// origin: languages/php/tests/php/test_php_extension_funcs_loaded_list.rs

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

$funcs = get_extension_funcs('non_existent_extension_xyz');
echo $funcs === false ? 'false_ok' : 'err', "\n";

__vybe_check(ob_get_clean(), "false_ok");
