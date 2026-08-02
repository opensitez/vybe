<?php
// vybe-test: php/oop_interfaces/interface_in_array_sum_via_mapped_typecheck
// origin: languages/php/tests/php/test_oop_interfaces.rs

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

interface Priced {
    public function price(): float;
}
class Product implements Priced { public function __construct(private float $p) {} public function price(): float { return $this->p; } }
$items = [new Product(1.2), new Product(2.8)];
$total = array_sum(array_map(fn(Priced $x) => $x->price(), $items));
echo $total;

__vybe_check(ob_get_clean(), "4");
