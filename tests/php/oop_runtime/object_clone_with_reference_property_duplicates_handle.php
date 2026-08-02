<?php
// vybe-test: php/oop_runtime/object_clone_with_reference_property_duplicates_handle
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

class Cnt {
    public function __construct(public array &$nums) {}
    public function value(): int { return $this->nums[0]; }
}
$arr = [1];
$a = new Cnt($arr);
$b = clone $a;
$arr[0] = 4;
echo $a->value() . $b->value();

__vybe_check(ob_get_clean(), "44");
