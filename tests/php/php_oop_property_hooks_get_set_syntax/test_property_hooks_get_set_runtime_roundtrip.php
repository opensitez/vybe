<?php
// vybe-test: php/php_oop_property_hooks_get_set_syntax/test_property_hooks_get_set_runtime_roundtrip
// origin: languages/php/tests/php/test_php_oop_property_hooks_get_set_syntax.rs

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

class Product {
    private float $_price = 0.0;

    public float $price {
        get => $this->_price;
        set => $this->_price = $value;
    }
}
$p = new Product();
$p->price = 19.95;
echo number_format($p->price, 2);

__vybe_check(ob_get_clean(), "19.95");
