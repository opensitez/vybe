<?php
// vybe-test: php/string_builtins_runtime/htmlspecialchars_disables_double_encode
// origin: languages/php/tests/php/test_string_builtins_runtime.rs

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

echo htmlspecialchars('<b>safe</b>', ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8', false);

__vybe_check(ob_get_clean(), "&lt;b&gt;safe&lt;/b&gt;");
