<?php
// vybe-test: php/patterns/value_object_equality_by_value
// origin: languages/php/tests/php/test_patterns.rs

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
    public function __construct(private int $amount, private string $currency) {}
    public function equals(Money $other): bool {
        return $this->amount === $other->amount && $this->currency === $other->currency;
    }
    public function add(Money $other): Money {
        if ($this->currency !== $other->currency) throw new \Exception('currency mismatch');
        return new Money($this->amount + $other->amount, $this->currency);
    }
    public function __toString(): string { return $this->amount . ' ' . $this->currency; }
}
$a = new Money(100, 'USD');
$b = new Money(100, 'USD');
$c = new Money(50, 'USD');
echo $a->equals($b) ? 'equal' : 'diff';
echo $a->equals($c) ? 'equal' : 'diff';
echo $a->add($c);

__vybe_check(ob_get_clean(), "equaldiff150 USD");
