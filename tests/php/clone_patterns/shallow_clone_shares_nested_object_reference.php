<?php
// vybe-test: php/clone_patterns/shallow_clone_shares_nested_object_reference
// origin: languages/php/tests/php/test_clone_patterns.rs

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

class Inner { public int $val = 0; }
class Outer { public Inner $inner; public function __construct() { $this->inner = new Inner(); } }
$a = new Outer();
$a->inner->val = 10;
$b = clone $a;
$b->inner->val = 99;
echo $a->inner->val;

__vybe_check(ob_get_clean(), "99");
