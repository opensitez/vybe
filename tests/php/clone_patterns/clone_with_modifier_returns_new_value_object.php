<?php
// vybe-test: php/clone_patterns/clone_with_modifier_returns_new_value_object
// origin: languages/php/tests/php/test_clone_patterns.rs

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
    public function __construct(private int $cents, private string $currency) {}
    public function add(int $cents): self {
        $new = clone $this;
        $new = new self($this->cents + $cents, $this->currency);
        return $new;
    }
    public function amount(): int { return $this->cents; }
    public function currency(): string { return $this->currency; }
}
$price = new Money(1000, 'USD');
$total = $price->add(500);
echo $price->amount() . ',' . $total->amount();

__vybe_check(ob_get_clean(), "1000,1500");
