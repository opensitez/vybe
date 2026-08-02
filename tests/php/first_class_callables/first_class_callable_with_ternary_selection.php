<?php
// vybe-test: php/first_class_callables/first_class_callable_with_ternary_selection
// origin: languages/php/tests/php/test_first_class_callables.rs

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

$select = true;
$fn = $select ? trim(... ) : strtoupper(...);
echo $fn('  php  ') . '|';
$select = false;
$fn = $select ? trim(... ) : strtoupper(...);
echo $fn('  php  ');

__vybe_check(ob_get_clean(), "php|PHP");
