<?php
// vybe-test: php/try_catch_nested_handlers/outer_try_inner_try_outer_catch_only_on_inner_throw
// origin: languages/php/tests/php/test_try_catch_nested_handlers.rs

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

$log = [];
try {
    try { $log[] = 'run'; throw new Exception('x'); }
    catch (Exception $e) { $log[] = 'inner'; }
} catch (Exception $e) {
    $log[] = 'outer';
}
echo implode(',', $log);

__vybe_check(ob_get_clean(), "run,inner");
