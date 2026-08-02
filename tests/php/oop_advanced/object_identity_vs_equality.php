<?php
// vybe-test: php/oop_advanced/object_identity_vs_equality
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

class Box {
    public function __construct(public int $value) {}
}
$a = new Box(5);
$b = $a;
$c = new Box(5);
echo ($a === $b) ? "same" : "different", "\n";
echo ($a === $c) ? "same" : "different", "\n";
echo ($a == $c) ? "equal" : "not equal", "\n";

__vybe_check(ob_get_clean(), "same\ndifferent\nequal");
