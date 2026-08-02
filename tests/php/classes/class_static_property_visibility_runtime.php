<?php
// vybe-test: php/classes/class_static_property_visibility_runtime
// origin: languages/php/tests/php/test_classes.rs

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
    public static function inc(): void { self::$count += 1; }
    public static function value(): int { return self::$count; }
}
Counter::inc();
Counter::inc();
echo Counter::value();

__vybe_check(ob_get_clean(), "2");
