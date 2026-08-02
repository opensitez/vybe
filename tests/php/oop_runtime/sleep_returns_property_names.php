<?php
// vybe-test: php/oop_runtime/sleep_returns_property_names
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

class S {
    public int $a = 1;
    private int $b = 2;
    public function __sleep(): array { return ['a']; }
}
$s = serialize(new S());
echo str_contains($s, 'a') && !str_contains($s, 'b') ? 'trim' : 'full';

__vybe_check(ob_get_clean(), "trim");
