<?php
// vybe-test: php/callables/is_callable_syntax_only_vs_runtime_exists
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

echo (is_callable('strlen', false) ? 'syn' : 'no')
   . ':'
   . (is_callable('strlen', true) ? 'rt' : 'no')
   . ':'
   . (is_callable('ghost_fn', false) ? 'syn' : 'no')
   . ':'
   . (is_callable('ghost_fn', true) ? 'rt' : 'no');

__vybe_check(ob_get_clean(), "syn:rt:no:rt");
