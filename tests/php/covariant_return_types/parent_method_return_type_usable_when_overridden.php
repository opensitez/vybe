<?php
// vybe-test: php/covariant_return_types/parent_method_return_type_usable_when_overridden
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

class A { public function value(): int { return 1; } }
class B extends A { public function value(): int { return parent::value() + 10; } }
echo (new B())->value();

__vybe_check(ob_get_clean(), "11");
