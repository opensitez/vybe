<?php
// vybe-test: php/late_static_binding/lsb_static_property_per_subclass
// origin: languages/php/tests/php/test_late_static_binding.rs

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

class Counter {
    protected static int $count = 0;
    public static function inc(): void { static::$count++; }
    public static function get(): int { return static::$count; }
}
class A extends Counter {}
class B extends Counter {}
A::inc(); A::inc(); B::inc();
echo A::get() . ',' . B::get();

__vybe_check(ob_get_clean(), "3,3");
