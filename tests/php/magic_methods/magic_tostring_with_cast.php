<?php
// vybe-test: php/magic_methods/magic_tostring_with_cast
// origin: languages/php/tests/php/test_magic_methods.rs

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

class Amount {
    public function __construct(private float $value) {}
    public function __toString(): string {
        return number_format($this->value, 2);
    }
}
$a = new Amount(42.5);
echo (string) $a;

__vybe_check(ob_get_clean(), "42.50");
