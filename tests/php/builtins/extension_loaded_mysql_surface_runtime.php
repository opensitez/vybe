<?php
// vybe-test: php/builtins/extension_loaded_mysql_surface_runtime
// origin: languages/php/tests/php/test_builtins.rs

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

echo extension_loaded('mysqlnd') ? 'yes' : 'no', "\n"; echo extension_loaded('mysqli') ? 'yes' : 'no', "\n"; echo extension_loaded('pdo_mysql') ? 'yes' : 'no', "\n"; echo extension_loaded('definitely_missing_ext') ? 'yes' : 'no', "\n";

__vybe_check(ob_get_clean(), "yes\nyes\nyes\nno");
