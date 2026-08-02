<?php
// vybe-test: php/namespaces/namespace_variable_function_name_resolves_global_when_qualified
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

function marker(string $v): string { return 'global-' . $v; }
namespace App {
    function marker(string $v): string { return 'local-' . $v; }
    $fn = '\\\\marker';
    echo $fn('x');
    echo '|';
    echo function_exists($fn) ? 'exists' : 'missing';
}

__vybe_check(ob_get_clean(), "global-x|exists");
