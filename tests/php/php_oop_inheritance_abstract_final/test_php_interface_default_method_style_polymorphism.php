<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_interface_default_method_style_polymorphism
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs

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

interface Sink {
    public function label(): string;
}
class A implements Sink {
    public function label(): string { return 'A'; }
}
class B implements Sink {
    public function label(): string { return 'B'; }
}
$xs = [new A(), new B()];
echo $xs[0]->label() . $xs[1]->label();

__vybe_check(ob_get_clean(), "AB");
