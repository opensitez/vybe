<?php
// vybe-test: php/oop_advanced/covariant_return
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

class Collection {
    protected array $items;
    public function __construct(array $items) { $this->items = $items; }
    public function first(): mixed { return $this->items[0] ?? null; }
}
class TypedCollection extends Collection {
    public function first(): string {
        return (string) parent::first();
    }
}
$c = new TypedCollection(["hello", "world"]);
echo $c->first(), "\n";

__vybe_check(ob_get_clean(), "hello");
