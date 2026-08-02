<?php
// vybe-test: php/covariant_return_types/interface_return_type_enforced_in_impl
// origin: languages/php/tests/php/test_covariant_return_types.rs

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

interface Transformer { public function transform(string $s): string; }
class UpperTransformer implements Transformer {
    public function transform(string $s): string { return strtoupper($s); }
}
$t = new UpperTransformer();
echo $t->transform("hello");

__vybe_check(ob_get_clean(), "HELLO");
