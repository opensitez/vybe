<?php
// vybe-test: php/magic_methods/magic_tostring_and_invoke_on_same_class
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

class Expression {
    public function __construct(private string $expr, private float $value) {}
    public function __toString(): string { return $this->expr . " = " . $this->value; }
    public function __invoke(float $factor): float { return $this->value * $factor; }
}
$e = new Expression("2+3", 5.0);
echo $e;
echo $e(3);

__vybe_check(ob_get_clean(), "2+3 = 515");
