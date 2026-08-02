<?php
// vybe-test: php/named_arguments_runtime/named_args_attribute_on_method
// origin: languages/php/tests/php/test_named_arguments_runtime.rs

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

class A {
    #[\Override]
    public function run(): string { return 'ok'; }
}
class B extends A { #[\Override] public function run(): string { return parent::run(); } }
echo (new B())->run();

__vybe_check(ob_get_clean(), "ok");
