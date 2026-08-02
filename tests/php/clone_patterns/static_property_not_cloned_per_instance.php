<?php
// vybe-test: php/clone_patterns/static_property_not_cloned_per_instance
// origin: languages/php/tests/php/test_clone_patterns.rs

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

class Counter { public static int $total = 0; public int $id; public function __construct() { $this->id = ++self::$total; } }
$a = new Counter();
$b = clone $a;
echo Counter::$total . ',' . $b->id;

__vybe_check(ob_get_clean(), "1,1");
