<?php
// vybe-test: php/php_streams_file_system_wrapper_ops/test_php_file_append_and_lock_flags
// origin: languages/php/tests/php/test_php_streams_file_system_wrapper_ops.rs

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

$tmp = tempnam(sys_get_temp_dir(), "vybe_append_");
file_put_contents($tmp, "Line 1\n");
file_put_contents($tmp, "Line 2\n", FILE_APPEND | LOCK_EX);
echo implode("-", file($tmp, FILE_IGNORE_NEW_LINES));
unlink($tmp);

__vybe_check(ob_get_clean(), "Line 1-Line 2");
