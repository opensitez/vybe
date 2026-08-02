<?php
// vybe-test: php/namespaces/namespace_function_call_resolution_with_use_as_alias_chain
// origin: languages/php/tests/php/test_namespaces.rs

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

namespace Runtime {
    function normalize(string $value): string { return strtoupper($value); }
}
namespace App {
    use function Runtime\normalize as up;
    echo up('ok');
}

__vybe_check(ob_get_clean(), "OK");
