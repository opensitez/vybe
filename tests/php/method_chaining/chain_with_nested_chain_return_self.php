<?php
// vybe-test: php/method_chaining/chain_with_nested_chain_return_self
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

class Logger {
    private array $events = [];
    public function push(string $e): static { $this->events[] = $e; return $this; }
    public function child(): static { return $this; }
    public function count(): int { return count($this->events); }
}
$log = new Logger();
echo $log->push('a')->push('b')->child()->push('c')->count();

__vybe_check(ob_get_clean(), "3");
