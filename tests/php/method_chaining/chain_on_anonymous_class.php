<?php
// vybe-test: php/method_chaining/chain_on_anonymous_class
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

$builder = new class {
    private array $items = [];
    public function push(string $v): static { $this->items[] = $v; return $this; }
    public function join(string $sep): string { return implode($sep, $this->items); }
};
echo $builder->push('x')->push('y')->join('-');

__vybe_check(ob_get_clean(), "x-y");
