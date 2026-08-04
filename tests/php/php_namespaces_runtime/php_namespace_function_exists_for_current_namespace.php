<?php
// vybe-test: php/php_namespaces_runtime/php_namespace_function_exists_for_current_namespace
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

namespace Local {
    function active(): bool { return true; }
}

namespace {
    use function Local\active;
    echo function_exists('active') ? 'no' : 'yes';
    echo '|';
    echo is_callable('Local\\active') ? 'callable' : 'nocal';
}

__vybe_check(ob_get_clean(), "yes|callable");
