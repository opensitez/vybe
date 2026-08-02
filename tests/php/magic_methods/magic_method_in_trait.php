<?php
// vybe-test: php/magic_methods/magic_method_in_trait
// origin: languages/php/tests/php/test_magic_methods.rs

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

trait DynamicAttributes {
    private array $attrs = [];
    public function __get($k) { return $this->attrs[$k] ?? null; }
    public function __set($k, $v) { $this->attrs[$k] = $v; }
    public function __isset($k) { return isset($this->attrs[$k]); }
}
class Product {
    use DynamicAttributes;
    public function __construct(public string $name) {}
}
$p = new Product("Widget");
$p->price = 9.99;
$p->stock = 100;
echo $p->name;
echo $p->price;
echo isset($p->stock) ? "in stock" : "out";

__vybe_check(ob_get_clean(), "Widget9.99in stock");
