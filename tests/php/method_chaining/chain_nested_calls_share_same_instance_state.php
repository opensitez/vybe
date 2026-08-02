<?php
// vybe-test: php/method_chaining/chain_nested_calls_share_same_instance_state
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

class Stack {
    private array $s = [];
    public function push(int $v): static { $this->s[] = $v; return $this; }
    public function top(): int { return $this->s[count($this->s) - 1]; }
}
$st = new Stack();
$st->push(1)->push(2);
echo $st->top();

__vybe_check(ob_get_clean(), "2");
