<?php
// vybe-test: php/named_arguments_runtime/named_args_constructor_promotion
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

class P { public function __construct(public string $name, public int $age) {} }
$p = new P(age: 30, name: 'Ann');
echo $p->name . $p->age;

__vybe_check(ob_get_clean(), "Ann30");
