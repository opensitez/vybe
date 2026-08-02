<?php
// vybe-test: php/type_juggling_runtime/is_bool_and_bool_as_string
// origin: languages/php/tests/php/test_type_juggling_runtime.rs

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

echo is_bool('true') ? 't1' : 't0';
echo '|';
echo is_bool((bool)'false') ? 'c1' : 'c0';
echo '|';
echo ((bool)0 === false) ? 'e1' : 'e0';
echo '|';
echo ((bool)'0' === false) ? 'f1' : 'f0';

__vybe_check(ob_get_clean(), "t0|c1|e1|f1");
