<?php
// vybe-test: php/late_static_binding/lsb_abstract_factory
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

abstract class Shape {
    abstract protected function area(): float;
    public static function describe(): string { return static::class . ' is a shape'; }
}
class Circle extends Shape { protected function area(): float { return 3.14; } }
echo Circle::describe();

__vybe_check(ob_get_clean(), "Circle is a shape");
