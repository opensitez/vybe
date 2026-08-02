<?php
// vybe-test: php/string_advanced/crc32_deterministic
// origin: languages/php/tests/php/test_string_advanced.rs

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

$a = crc32("hello");
$b = crc32("hello");
echo ($a === $b) ? "same" : "diff";
echo "\n";
echo ($a !== crc32("world")) ? "unique" : "collision";
echo "\n";

__vybe_check(ob_get_clean(), "same\nunique");
