<?php
// vybe-test: php/throw_expression_contexts/throw_in_serialize_allowed_classes_guard
// origin: languages/php/tests/php/test_throw_expression_contexts.rs

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

class Secret {}
$blob = serialize(new Secret());
try {
    unserialize($blob, ['allowed_classes' => false ? [Secret::class] : throw new RuntimeException('deny')]);
} catch (RuntimeException $e) { echo $e->getMessage(); }

__vybe_check(ob_get_clean(), "deny");
