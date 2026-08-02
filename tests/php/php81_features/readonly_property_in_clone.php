<?php
// vybe-test: php/php81_features/readonly_property_in_clone
// origin: languages/php/tests/php/test_php81_features.rs

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

class VO { public function __construct(public readonly int $val) {} }
$a = new VO(1);
$b = clone $a;
echo $a->val . ',' . $b->val;

__vybe_check(ob_get_clean(), "1,1");
