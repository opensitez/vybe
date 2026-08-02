<?php
// vybe-test: php/anonymous_classes/anon_class_inner_class_pattern
// origin: languages/php/tests/php/test_anonymous_classes.rs

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

class Outer {
    public function inner(): object {
        return new class($this) {
            public function __construct(private Outer $o) {}
            public function tag(): string { return 'inner-of-' . get_class($this->o); }
        };
    }
}
echo (new Outer)->inner()->tag();

__vybe_check(ob_get_clean(), "inner-of-Outer");
