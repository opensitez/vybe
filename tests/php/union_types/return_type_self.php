<?php
// vybe-test: php/union_types/return_type_self
// origin: languages/php/tests/php/test_union_types.rs

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

class Builder {
    private array $parts = [];
    public function add(string $p): self { $this->parts[] = $p; return $this; }
    public function build(): string { return implode('-', $this->parts); }
}
echo (new Builder)->add('a')->add('b')->build();

__vybe_check(ob_get_clean(), "a-b");
