<?php
// vybe-test: php/php_file_path_resolution_canonicalization/test_php_file_permissions_readable_writable_executable
// origin: languages/php/tests/php/test_php_file_path_resolution_canonicalization.rs

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

$tmpFile = tempnam(sys_get_temp_dir(), "vybe_perm_");
file_put_contents($tmpFile, "test perms");

echo is_readable($tmpFile) ? "R1" : "R0";
echo is_writable($tmpFile) ? "W1" : "W0";
echo is_file($tmpFile) ? "F1" : "F0";

unlink($tmpFile);

__vybe_check(ob_get_clean(), "R1W1F1");
