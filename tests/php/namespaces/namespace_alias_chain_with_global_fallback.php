<?php
// vybe-test: php/namespaces/namespace_alias_chain_with_global_fallback
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

namespace {
    function marker(string $v): string { return 'global:' . $v; }
}
namespace Package {
    function marker(string $v): string { return 'local:' . $v; }
}
namespace App {
    use function Package\\marker as local_marker;
    echo local_marker('x') . '|' . \Package\marker('y');
}

__vybe_check(ob_get_clean(), "local:x|local:y");
