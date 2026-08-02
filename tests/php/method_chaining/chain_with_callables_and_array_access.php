<?php
// vybe-test: php/method_chaining/chain_with_callables_and_array_access
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

class Bag {
    public array $items = [];
    public function push(string $v): static { $this->items[] = $v; return $this; }
}
$bag = (new Bag())->push('a')->push('b');
$bag->items[] = 'c';
echo implode('-', $bag->items);

__vybe_check(ob_get_clean(), "a-b-c");
