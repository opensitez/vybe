<?php
// vybe-test: php/php_cli_arguments_environment_vars/test_php_php_sapi_name_cli_detection
// origin: languages/php/tests/php/test_php_cli_arguments_environment_vars.rs

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

$sapi = php_sapi_name();
echo (strlen($sapi) > 0) ? "SAPI_AVAILABLE" : "NO_SAPI";

__vybe_check(ob_get_clean(), "SAPI_AVAILABLE");
