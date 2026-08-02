<?php
// vybe-test: php/method_chaining/chain_with_method_returning_array
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

class Collector {
    private array $items = [];
    public function add(string $v): static { $this->items[] = $v; return $this; }
    public function all(): array { return $this->items; }
}
echo implode('-', (new Collector())->add('p')->add('q')->all());

__vybe_check(ob_get_clean(), "p-q");
