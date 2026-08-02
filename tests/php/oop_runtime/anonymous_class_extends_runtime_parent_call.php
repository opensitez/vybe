<?php
// vybe-test: php/oop_runtime/anonymous_class_extends_runtime_parent_call
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

class Base {
    public function label(): string { return 'base'; }
}
$o = new class extends Base {
    public function label(): string { return parent::label() . '+anon'; }
};
echo $o->label();

__vybe_check(ob_get_clean(), "base+anon");
