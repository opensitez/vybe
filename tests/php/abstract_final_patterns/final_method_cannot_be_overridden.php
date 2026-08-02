<?php
// vybe-test: php/abstract_final_patterns/final_method_cannot_be_overridden
// origin: languages/php/tests/php/test_abstract_final_patterns.rs

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

class Base {
    final public function version(): string { return "1.0"; }
}
try {
    eval('class Child extends Base { public function version(): string { return "2.0"; } }');
} catch (\Error $e) {
    echo "cannot override", "\n";
}

__vybe_check(ob_get_clean(), "cannot override");
