<?php
// vybe-test: php/static_members_runtime/static_promoted_readonly_not_applicable_use_normal
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

class Point { public static int $x = 0; public static function set(int $v): void { self::$x = $v; } }
Point::set(4);
echo Point::$x;

__vybe_check(ob_get_clean(), "4");
