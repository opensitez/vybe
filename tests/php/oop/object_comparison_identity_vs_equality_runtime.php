<?php
// vybe-test: php/oop/object_comparison_identity_vs_equality_runtime
// origin: languages/php/tests/php/test_oop.rs

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

class Marker { public function __construct(public int $v) {} }
$a = new Marker(1);
$b = new Marker(1);
$c = $a;
echo ($a == $b) ? 'eq' : 'neq';
echo ($a === $c) ? '|same' : '|diff';

__vybe_check(ob_get_clean(), "eq|same");
