<?php
// vybe-test: php/php_oop_late_static_binding_self_static/test_php_static_property_shadowing_with_parent_access
// origin: languages/php/tests/php/test_php_oop_late_static_binding_self_static.rs

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

class BaseCounter {
    protected static int $count = 1;
    public static function current(): int {
        return static::$count;
    }
}

class DerivedCounter extends BaseCounter {
    protected static int $count = 4;
    public static function currentFromParent(): int {
        return parent::$count;
    }
}

echo DerivedCounter::current() . "|" . DerivedCounter::currentFromParent();

__vybe_check(ob_get_clean(), "4|1");
