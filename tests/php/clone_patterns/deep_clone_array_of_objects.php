<?php
// vybe-test: php/clone_patterns/deep_clone_array_of_objects
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

class Item { public function __construct(public int $id) {} }
class Cart {
    public array $items = [];
    public function __clone() {
        $this->items = array_map(fn($i) => clone $i, $this->items);
    }
}
$cart = new Cart();
$cart->items[] = new Item(1);
$cart->items[] = new Item(2);
$copy = clone $cart;
$copy->items[0]->id = 99;
echo $cart->items[0]->id . ',' . $copy->items[0]->id;

__vybe_check(ob_get_clean(), "1,99");
