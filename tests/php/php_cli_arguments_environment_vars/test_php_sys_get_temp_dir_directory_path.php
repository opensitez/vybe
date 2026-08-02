<?php
// vybe-test: php/php_cli_arguments_environment_vars/test_php_sys_get_temp_dir_directory_path
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

$tmp = sys_get_temp_dir();
echo is_dir($tmp) ? "TEMP_DIR_EXISTS" : "TEMP_DIR_MISSING";

__vybe_check(ob_get_clean(), "TEMP_DIR_EXISTS");
