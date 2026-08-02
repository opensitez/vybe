<?php
// vybe-test: php/type_hints_advanced/nullsafe_null_returns_null_not_error
// origin: languages/php/tests/php/test_type_hints_advanced.rs

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

class A { public function b(): ?B { return null; } }
class B { public function c(): string { return 'c'; } }
$result = (new A)->b()?->c();
echo $result ?? 'null';

__vybe_check(ob_get_clean(), "null");
