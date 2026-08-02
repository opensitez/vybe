<?php
// vybe-test: php/static_members_runtime/static_private_accessible_in_class
// origin: languages/php/tests/php/test_static_members_runtime.rs

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

class Vault { private static int $secret = 9; public static function peek(): int { return self::$secret; } }
echo Vault::peek();

__vybe_check(ob_get_clean(), "9");
