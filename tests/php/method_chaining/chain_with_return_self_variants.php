<?php
// vybe-test: php/method_chaining/chain_with_return_self_variants
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

class Builder {
    private array $steps = [];
    public function set(string $name, int $value): static { $this->steps[$name] = $value; return $this; }
    public function merge(array $data): static { $this->steps = array_merge($this->steps, $data); return $this; }
    public function size(): int { return count($this->steps); }
}
echo (new Builder())
    ->set('a', 1)
    ->merge(['b' => 2])
    ->set('c', 3)
    ->size();

__vybe_check(ob_get_clean(), "3");
