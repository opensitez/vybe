<?php
// vybe-test: php/namespaces/namespace_local_and_global_function_resolution
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

function marker(): string { return 'global-marker'; }
namespace App {
    function marker(): string { return 'local-marker'; }
    echo marker();
    echo '|';
    echo \marker();
}

__vybe_check(ob_get_clean(), "local-marker|global-marker");
