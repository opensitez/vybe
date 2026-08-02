<?php
// vybe-test: php/advanced_oop/static_property_per_class_hierarchy
// origin: languages/php/tests/php/test_advanced_oop.rs

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

class CounterA {
    public static int $count = 1;
    public static function bump(): void { self::$count += 1; }
}
class CounterB {
    public static int $count = 10;
    public static function bump(): void { self::$count += 10; }
}
CounterA::bump();
CounterB::bump();
echo CounterA::$count . '|' . CounterB::$count;

__vybe_check(ob_get_clean(), "2|20");
