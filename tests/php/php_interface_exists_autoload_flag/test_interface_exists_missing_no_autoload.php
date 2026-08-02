<?php
// vybe-test: php/php_interface_exists_autoload_flag/test_interface_exists_missing_no_autoload
// origin: languages/php/tests/php/test_php_interface_exists_autoload_flag.rs

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

echo interface_exists('NonExistentContract', false) ? 'found' : 'missing_no_autoload', "\n";

__vybe_check(ob_get_clean(), "missing_no_autoload");
