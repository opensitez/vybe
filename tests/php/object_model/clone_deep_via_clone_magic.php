<?php
// vybe-test: php/object_model/clone_deep_via_clone_magic
// origin: languages/php/tests/php/test_object_model.rs

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

class Inner { public int $v = 0; }
class Outer {
    public Inner $inner;
    public function __construct() { $this->inner = new Inner; }
    public function __clone() { $this->inner = clone $this->inner; }
}
$a = new Outer; $a->inner->v = 5;
$b = clone $a; $b->inner->v = 99;
echo $a->inner->v . ',' . $b->inner->v;

__vybe_check(ob_get_clean(), "5,99");
