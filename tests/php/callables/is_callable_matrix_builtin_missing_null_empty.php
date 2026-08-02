<?php
// vybe-test: php/callables/is_callable_matrix_builtin_missing_null_empty
// origin: languages/php/tests/php/test_callables.rs

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

echo (is_callable('strlen') ? 'S' : '-')
   . (is_callable('missing_fn_xyz') ? 'S' : 'M')
   . (is_callable(null) ? 'S' : 'N')
   . (is_callable('') ? 'S' : 'E');

__vybe_check(ob_get_clean(), "SMNE");
