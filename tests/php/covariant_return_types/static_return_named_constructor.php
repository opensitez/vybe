<?php
// vybe-test: php/covariant_return_types/static_return_named_constructor
// origin: languages/php/tests/php/test_covariant_return_types.rs

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
    private function __construct(private int $cents) {}
    public static function fromCents(int $cents): static { return new static($cents); }
    public function amount(): int { return $this->cents; }
}
class Euro extends Money {}
$e = Euro::fromCents(500);
echo $e->amount();

__vybe_check(ob_get_clean(), "500");
