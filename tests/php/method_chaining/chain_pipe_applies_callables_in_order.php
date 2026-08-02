<?php
// vybe-test: php/method_chaining/chain_pipe_applies_callables_in_order
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

class Pipe {
    public function __construct(private mixed $value) {}
    public function through(callable $fn): static {
        $this->value = $fn($this->value);
        return $this;
    }
    public function out(): mixed { return $this->value; }
}
echo (new Pipe(4))->through(fn($n) => $n * 2)->through(fn($n) => $n + 1)->out();

__vybe_check(ob_get_clean(), "9");
