<?php
// vybe-test: php/covariant_return_types/self_return_type_stays_in_declared_class
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

class Base {
    public function withData(string $d): self {
        $clone = clone $this;
        return $clone;
    }
    public function type(): string { return "Base"; }
}
$b = new Base();
echo $b->withData("x")->type();

__vybe_check(ob_get_clean(), "Base");
