<?php
// vybe-test: php/try_catch_nested_handlers/catch_variable_scope_not_visible_after_block
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

try { throw new Exception('scoped'); }
catch (Exception $e) { $inside = $e->getMessage(); }
finally { echo isset($inside) ? $inside : 'gone'; }

__vybe_check(ob_get_clean(), "scoped");
