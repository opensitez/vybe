<?php
// vybe-test: php/method_chaining/chain_through_array_map_closure
// origin: languages/php/tests/php/test_method_chaining.rs

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

class ListBuilder {
    private array $items = [];
    public function add(int $v): static { $this->items[] = $v; return $this; }
    public function map(callable $fn): static { $this->items = array_map($fn, $this->items); return $this; }
    public function first(): int { return $this->items[0]; }
}
echo (new ListBuilder())->add(1)->add(2)->map(fn($n) => $n * 10)->first();

__vybe_check(ob_get_clean(), "10");
