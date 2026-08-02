<?php
// vybe-test: php/classes/class_implements_multiple_interfaces_runtime
// origin: languages/php/tests/php/test_classes.rs

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

interface A { public function a(): string; }
interface B { public function b(): string; }
class Impl implements A, B {
    public function a(): string { return 'A'; }
    public function b(): string { return 'B'; }
}
echo (new Impl()) instanceof A ? 'A' : 'no';
echo '|';
echo (new Impl()) instanceof B ? 'B' : 'no';

__vybe_check(ob_get_clean(), "A|B");
