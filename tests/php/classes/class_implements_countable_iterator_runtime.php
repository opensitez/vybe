<?php
// vybe-test: php/classes/class_implements_countable_iterator_runtime
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

class Bag implements Countable {
    public function __construct(private array $items) {}
    public function count(): int { return count($this->items); }
}
echo (new Bag([1, 2, 3, 4]))->count();

__vybe_check(ob_get_clean(), "4");
