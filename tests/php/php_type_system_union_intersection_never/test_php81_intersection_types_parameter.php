<?php
// vybe-test: php/php_type_system_union_intersection_never/test_php81_intersection_types_parameter
// origin: languages/php/tests/php/test_php_type_system_union_intersection_never.rs

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

interface CountableCollection extends Countable, ArrayAccess {}

class CustomCollection implements CountableCollection {
    private array $items = [10, 20];
    public function count(): int { return count($this->items); }
    public function offsetExists($o): bool { return isset($this->items[$o]); }
    public function offsetGet($o): mixed { return $this->items[$o]; }
    public function offsetSet($o, $v): void {}
    public function offsetUnset($o): void {}
}

function inspect(Countable&ArrayAccess $coll): string {
    return "Count=" . count($coll) . " First=" . $coll[0];
}

echo inspect(new CustomCollection());

__vybe_check(ob_get_clean(), "Count=2 First=10");
