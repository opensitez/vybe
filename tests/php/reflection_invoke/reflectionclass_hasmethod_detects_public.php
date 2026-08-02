<?php
// vybe-test: php/reflection_invoke/reflectionclass_hasmethod_detects_public
// origin: languages/php/tests/php/test_reflection_invoke.rs

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

class S { public function go(): void {} }
$ref = new ReflectionClass(S::class);
echo $ref->hasMethod('go') ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes");
