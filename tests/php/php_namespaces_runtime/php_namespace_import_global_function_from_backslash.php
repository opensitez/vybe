<?php
// vybe-test: php/php_namespaces_runtime/php_namespace_import_global_function_from_backslash
// origin: languages/php/tests/php/test_php_namespaces_runtime.rs

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

namespace Demo;
function trim_global(string $value): string { return 'local'; }

namespace {
    use function Demo\trim_global as local_trim;
    echo local_trim('x') . '|' . \trim(' x ');
}

__vybe_check(ob_get_clean(), "local|x");
