<?php
// vybe-test: php/php_magic_methods_get_set_call_invoke/test_php_magic_to_string_cast
// origin: languages/php/tests/php/test_php_magic_methods_get_set_call_invoke.rs

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

class Money {
    public function __construct(public float $amount, public string $currency) {}
    public function __toString(): string {
        return "{$this->currency} " . number_format($this->amount, 2);
    }
}

$m = new Money(99.9, "USD");
echo "Total: $m";

__vybe_check(ob_get_clean(), "Total: USD 99.90");
