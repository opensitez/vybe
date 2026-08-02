<?php
// vybe-test: php/oop_runtime/clone_copies_private_property_values
// origin: languages/php/tests/php/test_oop_runtime.rs

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

class Secret {
    public function __construct(private int $n) {}
    public function get(): int { return $this->n; }
}
$a = new Secret(8);
$b = clone $a;
echo $a->get() . $b->get();

__vybe_check(ob_get_clean(), "88");
