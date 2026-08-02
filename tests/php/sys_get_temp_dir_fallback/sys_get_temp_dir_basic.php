<?php
// vybe-test: php/sys_get_temp_dir_fallback/sys_get_temp_dir_basic
// origin: languages/php/tests/php/test_sys_get_temp_dir_fallback.rs

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

$dir = sys_get_temp_dir();
echo is_string($dir) && strlen($dir) > 0 ? "ok" : "fail";

__vybe_check(ob_get_clean(), "ok");
