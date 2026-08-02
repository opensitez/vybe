<?php
// vybe-test: php/advanced_oop/named_constructor_via_interface
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

interface HasNamedConstructors {
    public static function empty(): static;
}
class Collection implements HasNamedConstructors {
    private function __construct(private array $items = []) {}
    public static function empty(): static { return new static; }
    public function count(): int { return count($this->items); }
}
echo Collection::empty()->count();

__vybe_check(ob_get_clean(), "0");
