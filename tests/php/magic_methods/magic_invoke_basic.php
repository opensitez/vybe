<?php
// vybe-test: php/magic_methods/magic_invoke_basic
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

class Multiplier {
    public function __construct(private int $factor) {}
    public function __invoke($x) {
        return $x * $this->factor;
    }
}
$double = new Multiplier(2);
echo $double(5);
echo $double(10);

__vybe_check(ob_get_clean(), "1020");
