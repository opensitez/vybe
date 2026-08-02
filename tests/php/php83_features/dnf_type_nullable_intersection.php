<?php
// vybe-test: php/php83_features/dnf_type_nullable_intersection
// origin: languages/php/tests/php/test_php83_features.rs

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

interface A {}
interface B {}
class C implements A, B {}
function test((A&B)|null $obj): string {
    return $obj === null ? 'null' : 'obj';
}
echo test(new C) . ',' . test(null);

__vybe_check(ob_get_clean(), "obj,null");
