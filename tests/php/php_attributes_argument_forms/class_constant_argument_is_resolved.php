<?php
// vybe-test: php/php_attributes_argument_forms/class_constant_argument_is_resolved
// origin: languages/php/tests/php/test_php_attributes_argument_forms.rs

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

class Limits {
    const MAX = 100;
}
#[Attribute]
class Cap {
    public function __construct(public int $n) {}
}
#[Cap(Limits::MAX)]
class Bucket {}
echo (new ReflectionClass(Bucket::class))->getAttributes(Cap::class)[0]->newInstance()->n;

__vybe_check(ob_get_clean(), "100");
