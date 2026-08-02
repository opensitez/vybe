<?php
// vybe-test: php/php_set_exception_handler_nested_throws/test_php_restore_exception_handler_reverts_to_previous
// origin: languages/php/tests/php/test_php_set_exception_handler_nested_throws.rs

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
$h1 = function(Throwable $e) use (&$log) { $log[] = "H1:" . $e->getMessage(); };
$h2 = function(Throwable $e) use (&$log) { $log[] = "H2:" . $e->getMessage(); };

set_exception_handler($h1);
set_exception_handler($h2);
restore_exception_handler(); // Reverts to H1

$current = set_exception_handler(null);
$current(new Exception("Event"));

echo implode(", ", $log);

__vybe_check(ob_get_clean(), "H1:Event");
