<?php
// vybe-test: php/modern_php_deep/readonly_via_constructor_promotion
// origin: languages/php/tests/php/test_modern_php_deep.rs

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
    public function __construct(
        public readonly int    $amount,
        public readonly string $currency
    ) {}
    public function format(): string {
        return $this->amount . " " . $this->currency;
    }
}
$m = new Money(100, "USD");
echo $m->amount;
echo $m->currency;
echo $m->format();

__vybe_check(ob_get_clean(), "100USD100 USD");
