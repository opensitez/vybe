<?php
// vybe-test: php/namespaces/namespace_relative_to_global_function_lookup_with_use
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

function global_marker(string $v): string { return 'g:' . $v; }
namespace App {
    use function global_marker as marker;
    function local(string $v): string { return marker($v); }
    echo local('x');
}

__vybe_check(ob_get_clean(), "g:x");
