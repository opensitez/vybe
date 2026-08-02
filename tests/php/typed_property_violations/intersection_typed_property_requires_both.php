<?php
// vybe-test: php/typed_property_violations/intersection_typed_property_requires_both
// origin: languages/php/tests/php/test_typed_property_violations.rs

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

interface A { public function a(): void; }
interface B { public function b(): void; }
class Both implements A, B { public function a(): void {} public function b(): void {} }
class Holder { public A&B $item; }
$h = new Holder();
$h->item = new Both();
echo $h->item instanceof Both ? 'both' : 'no';

__vybe_check(ob_get_clean(), "both");
