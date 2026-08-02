<?php
// vybe-test: php/design_patterns/value_object_immutability
// origin: languages/php/tests/php/test_design_patterns.rs

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

final class Money {
    public function __construct(private readonly int $amount, private readonly string $currency) {}
    public function add(Money $other): self {
        if ($this->currency !== $other->currency) throw new \InvalidArgumentException('Currency mismatch');
        return new self($this->amount + $other->amount, $this->currency);
    }
    public function __toString(): string { return $this->amount . ' ' . $this->currency; }
}
$a = new Money(100, 'USD');
$b = new Money(50, 'USD');
echo $a->add($b);

__vybe_check(ob_get_clean(), "150 USD");
