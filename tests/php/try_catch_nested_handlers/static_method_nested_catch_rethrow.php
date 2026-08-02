<?php
// vybe-test: php/try_catch_nested_handlers/static_method_nested_catch_rethrow
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

class Api {
    public static function call(): void {
        try {
            try { throw new RuntimeException('api'); }
            catch (RuntimeException $e) { throw new Exception('mapped'); }
        } catch (Exception $e) {
            echo $e->getMessage();
        }
    }
}
Api::call();

__vybe_check(ob_get_clean(), "mapped");
