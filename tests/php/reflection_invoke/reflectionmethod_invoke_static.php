<?php
// vybe-test: php/reflection_invoke/reflectionmethod_invoke_static
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

class Id { public static function val(): int { return 42; } }
$ref = new ReflectionMethod(Id::class, 'val');
echo $ref->invoke(null);

__vybe_check(ob_get_clean(), "42");
