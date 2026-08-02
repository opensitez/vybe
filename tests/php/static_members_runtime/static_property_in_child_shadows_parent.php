<?php
// vybe-test: php/static_members_runtime/static_property_in_child_shadows_parent
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

class Base { public static int $n = 1; }
class Child extends Base { public static int $n = 2; }
echo Child::$n;

__vybe_check(ob_get_clean(), "2");
