<?php
// vybe-test: php/php_bitwise_shifts_flags_masking/test_php_bitwise_permission_masking_pattern
// origin: languages/php/tests/php/test_php_bitwise_shifts_flags_masking.rs

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

const PERM_READ = 1 << 0;  // 1
const PERM_WRITE = 1 << 1; // 2
const PERM_EXEC = 1 << 2;  // 4

$userPerms = PERM_READ | PERM_EXEC;

echo ($userPerms & PERM_READ ? "1" : "0");
echo ($userPerms & PERM_WRITE ? "1" : "0");
echo ($userPerms & PERM_EXEC ? "1" : "0");

__vybe_check(ob_get_clean(), "101");
