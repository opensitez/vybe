<?php
// vybe-test: php/oop_advanced/objects_in_array_map
// origin: languages/php/tests/php/test_oop_advanced.rs

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
    public function __construct(public string $name, public float $price) {}
    public function discounted(float $pct): self {
        return new self($this->name, $this->price * (1 - $pct));
    }
}
$items = [
    new Item("Widget", 10.0),
    new Item("Gadget", 20.0),
    new Item("Doohickey", 5.0),
];
$discounted = array_map(fn($i) => $i->discounted(0.1), $items);
$totals = array_map(fn($i) => $i->price, $discounted);
echo number_format(array_sum($totals), 2), "\n";

__vybe_check(ob_get_clean(), "31.50");
