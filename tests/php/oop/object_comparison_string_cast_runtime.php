<?php
// vybe-test: php/oop/object_comparison_string_cast_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class Item {
    public string $id;
    public function __construct(string $id) { $this->id = $id; }
    public function __toString(): string { return $this->id; }
}
$one = new Item('A');
$two = new Item('A');
echo (string)$one;
echo (string)$two;
echo ($one == $two) ? '|eq' : '|neq';

__vybe_check(ob_get_clean(), "AA|eq");
