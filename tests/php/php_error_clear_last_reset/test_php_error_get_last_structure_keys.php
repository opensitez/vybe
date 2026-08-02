<?php
// vybe-test: php/php_error_clear_last_reset/test_php_error_get_last_structure_keys
// origin: languages/php/tests/php/test_php_error_clear_last_reset.rs

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

@trigger_error("Structured error", E_USER_WARNING);
$err = error_get_last();
error_clear_last();

echo "Type={$err['type']} Message={$err['message']}";

__vybe_check(ob_get_clean(), "Type=512 Message=Structured error");
