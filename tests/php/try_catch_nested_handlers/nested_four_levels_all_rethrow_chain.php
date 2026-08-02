<?php
// vybe-test: php/try_catch_nested_handlers/nested_four_levels_all_rethrow_chain
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
    try {
        try {
            try { throw new LogicException('4'); }
            catch (LogicException $e) { $log[] = 'c4'; throw $e; }
        } catch (LogicException $e) { $log[] = 'c3'; throw $e; }
    } catch (LogicException $e) { $log[] = 'c2'; throw $e; }
} catch (LogicException $e) { $log[] = 'c1'; }
echo implode(',', $log);

__vybe_check(ob_get_clean(), "c4,c3,c2,c1");
