<?php
// vybe-test: php/oop_advanced/clone_deep_copy_array_property
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

class ShoppingCart {
    private array $items = [];
    public function add(string $item): void {
        $this->items[] = $item;
    }
    public function __clone() {
        // array is value-type so no manual deep copy needed,
        // but let's verify mutation isolation
        $this->items = $this->items; // explicit reassign
    }
    public function count(): int { return count($this->items); }
    public function items(): array { return $this->items; }
}
$cart1 = new ShoppingCart();
$cart1->add("apple");
$cart1->add("banana");
$cart2 = clone $cart1;
$cart2->add("cherry");
echo $cart1->count(), "\n";
echo $cart2->count(), "\n";
echo implode(",", $cart1->items()), "\n";

__vybe_check(ob_get_clean(), "2\n3\napple,banana");
