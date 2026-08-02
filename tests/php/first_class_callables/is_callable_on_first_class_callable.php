<?php
// vybe-test: php/first_class_callables/is_callable_on_first_class_callable
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

$fn = strlen(...);
echo is_callable($fn) ? 'callable' : 'not callable';
echo "\n";
$staticFn = DateTime::createFromFormat(...);
echo is_callable($staticFn) ? 'callable' : 'not callable';
echo "\n";

__vybe_check(ob_get_clean(), "callable\ncallable");
