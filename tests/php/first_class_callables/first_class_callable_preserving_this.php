<?php
// vybe-test: php/first_class_callables/first_class_callable_preserving_this
// origin: languages/php/tests/php/test_first_class_callables.rs

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
    private int $count = 0;
    public function increment(int $by = 1): void { $this->count += $by; }
    public function getCount(): int { return $this->count; }
}
$counter = new Counter();
$inc = $counter->increment(...);
$get = $counter->getCount(...);
$inc(5);
$inc(3);
echo $get() . "\n";
array_map($inc, [1, 1, 1]);
echo $get() . "\n";

__vybe_check(ob_get_clean(), "8\n11");
