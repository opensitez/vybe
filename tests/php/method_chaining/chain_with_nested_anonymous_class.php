<?php
// vybe-test: php/method_chaining/chain_with_nested_anonymous_class
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
    private array $parts = [];
    public function add(string $s): static { $this->parts[] = $s; return $this; }
    public function chain(int $i): self {
        $this->parts[] = (string)$i;
        return $this;
    }
    public function join(): string { return implode('|', $this->parts); }
};
echo $builder->add('x')->chain(9)->join();

__vybe_check(ob_get_clean(), "x|9");
