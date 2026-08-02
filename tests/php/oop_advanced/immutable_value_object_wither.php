<?php
// vybe-test: php/oop_advanced/immutable_value_object_wither
// origin: languages/php/tests/php/test_oop_advanced.rs

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
        public readonly int $amount,
        public readonly string $currency,
    ) {}
    public function withAmount(int $amount): self {
        return new self($amount, $this->currency);
    }
    public function withCurrency(string $currency): self {
        return new self($this->amount, $currency);
    }
    public function __toString(): string {
        return "{$this->amount} {$this->currency}";
    }
}
$m1 = new Money(100, "USD");
$m2 = $m1->withAmount(200);
$m3 = $m2->withCurrency("EUR");
echo $m1, "\n";
echo $m2, "\n";
echo $m3, "\n";

__vybe_check(ob_get_clean(), "100 USD\n200 USD\n200 EUR");
