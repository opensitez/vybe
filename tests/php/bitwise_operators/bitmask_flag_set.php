<?php
// vybe-test: php/bitwise_operators/bitmask_flag_set
// origin: languages/php/tests/php/test_bitwise_operators.rs

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

define('READ',  0b001);
define('WRITE', 0b010);
define('EXEC',  0b100);
$perms = READ | WRITE;
echo ($perms & READ)  ? 'r' : '-';
echo ($perms & WRITE) ? 'w' : '-';
echo ($perms & EXEC)  ? 'x' : '-';

__vybe_check(ob_get_clean(), "rw-");
