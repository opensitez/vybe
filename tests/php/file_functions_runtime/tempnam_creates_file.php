<?php
// vybe-test: php/file_functions_runtime/tempnam_creates_file
// origin: languages/php/tests/php/test_file_functions_runtime.rs

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

$p = tempnam(sys_get_temp_dir(), 'vybe');
echo is_string($p) && file_exists($p) ? 'ok' : 'no';
unlink($p);

__vybe_check(ob_get_clean(), "ok");
