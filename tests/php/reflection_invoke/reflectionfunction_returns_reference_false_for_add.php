<?php
// vybe-test: php/reflection_invoke/reflectionfunction_returns_reference_false_for_add
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

function add(int $a, int $b): int { return $a + $b; }
$ref = new ReflectionFunction('add');
echo $ref->returnsReference() ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "no");
