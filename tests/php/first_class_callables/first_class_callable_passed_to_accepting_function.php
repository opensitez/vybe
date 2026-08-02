<?php
// vybe-test: php/first_class_callables/first_class_callable_passed_to_accepting_function
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

function applyAll(array $fns, $value) {
    return array_reduce($fns, fn($carry, $fn) => $fn($carry), $value);
}
$result = applyAll([strtoupper(...), trim(...), strrev(...)], '  hello  ');
echo $result . "\n";

__vybe_check(ob_get_clean(), "OLLEH");
