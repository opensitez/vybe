<?php
// vybe-test: php/oop_patterns/static_property_shared_state_runtime
// origin: languages/php/tests/php/test_oop_patterns.rs

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
    private static int $count = 0;
    public function __construct() { self::$count++; }
    public static function total(): int { return self::$count; }
}
new Counter();
new Counter();
echo (string) Counter::total();

__vybe_check(ob_get_clean(), "2");
