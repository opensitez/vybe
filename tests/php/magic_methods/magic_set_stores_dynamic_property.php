<?php
// vybe-test: php/magic_methods/magic_set_stores_dynamic_property
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

class Bag {
    private array $items = [];
    public function __set($k, $v) { $this->items[$k] = $v; }
    public function __get($k) { return $this->items[$k] ?? null; }
    public function keys(): array { return array_keys($this->items); }
}
$b = new Bag();
$b->x = 10;
$b->y = 20;
$b->z = 30;
echo implode(",", $b->keys());
echo $b->y;

__vybe_check(ob_get_clean(), "x,y,z20");
